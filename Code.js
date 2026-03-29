const FOLDER_ID = '1WoTjgUiF4PjVQNBX6chRQzDFUP2f7k3a';
const TARGET_HOUR = 9; // 9:00 AM corresponds to a .375 decimal part in Google Sheets

function onOpen()
{
	SpreadsheetApp.getUi()
		.createMenu('Conventions')
		.addItem('Importer les données Structures', 'collectStructuresFromIndividualFiles')
		.addItem('Importer les données Personnes', 'collectPersonnesFromIndividualFiles')
		.addSeparator()
		.addItem('Simuler la correction des erreurs (Dry Run)', 'dryRunFixClericalErrors')
		.addItem('Appliquer la correction des erreurs', 'applyFixClericalErrors')
		.addToUi();
}

/**
 * Normalizes a value if it's a Date by setting its time to 9:00 AM.
 * This ensures the numeric value in Google Sheets ends with a .375 decimal part.
 * 
 * @param {any} value The value to normalize.
 * @returns {any} The normalized value.
 */
function normalizeDate(value)
{
	if (value instanceof Date)
	{
		const normalized = new Date(value.getTime());
		normalized.setHours(TARGET_HOUR, 0, 0, 0);
		return normalized;
	}
	return value;
}

/**
 * Configuration for automatic corrections.
 * Each key is a cell address (Sheet!A1).
 * Each value is an object with a name and a 'fixer' function.
 *
 * @typedef {Object} FixResult
 * @property {boolean} [success] Indicates if the correction was successful.
 * @property {string|Date} [fixedValue] The corrected value.
 * @property {boolean} [modified] Indicates if the value was modified.
 * @property {string} [error] Error message if the correction failed.
 *
 * @type {Object<string, {name: string, fixer: (value: any, ss: GoogleAppsScript.Spreadsheet.Spreadsheet) => FixResult}>}
 */
const FIX_CONFIG = {
	'Saisie!C2': {
		name: 'Numéro de convention',
		fixer: (value, ss) =>
		{
			let strValue = String(value || ss?.getName() || '').trim();

			const regexp = /^[Bb]*0*([1-9][0-9]*)/;
			const match  = regexp.exec(strValue);

			if (match && match[1])
			{
				const fixedValue = match[1];

				return { success: true, fixedValue: fixedValue, modified: (fixedValue !== strValue) };
			}

			return { error: `La valeur "${strValue}" ne correspond pas à un format d'entier valide.` };
		}
	},
	'Saisie!C72': {
		name: 'Date de signature',
		fixer: (value, ss) =>
		{
			let strValue;
			let originalIsDate = false;
			let originalTimeCorrect = false;

			if (value instanceof Date)
			{
				originalIsDate = true;
				originalTimeCorrect = (value.getHours() === TARGET_HOUR && value.getMinutes() === 0 && value.getSeconds() === 0);
				strValue = Utilities.formatDate(value, Session.getScriptTimeZone(), 'dd/MM/yyyy');
			}
			else
			{
				strValue = String(value || '').trim();
			}

			if (strValue === '')
			{
				return { error: 'La valeur est vide.' };
			}

			const originalStrValue = strValue;

			// Attempt to correct if the format is D/MYYYY or DD/MMYYYY (missing the second slash)
			const dateRegex = /^([0-9]{1,2})\/([0-9]{1,2})([0-9]{4}|[0-9]{2})$/;
			const match = strValue.match(dateRegex);
			if (match)
			{
				strValue = `${match[1]}/${match[2]}/${match[3]}`;
			}

			// Date parsing
			const parts = strValue.split(/[\/\-\.]/);
			if (parts.length !== 3)
			{
				return { error: `Format de date invalide : "${strValue}". Attendu : DD/MM/YYYY.` };
			}

			let day = parseInt(parts[0], 10);
			let month = parseInt(parts[1], 10) - 1; // 0-indexed
			let year = parseInt(parts[2], 10);

			if (year < 100)
			{
				year += 2000;
			}

			const date = new Date(year, month, day, TARGET_HOUR, 0, 0, 0);
			const now = new Date;
			const minDate = new Date(2024, 0, 1);

			// Verification of actual date validity
			if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day)
			{
				return { error: `La date "${strValue}" est calendairement invalide.` };
			}

			if (date < minDate || date > now)
			{
				return { error: `La date "${strValue}" est hors limites (doit être entre 01/01/2024 et aujourd'hui).` };
			}

			const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
			const modified = !originalIsDate || !originalTimeCorrect || (formattedDate !== originalStrValue);

			return { success: true, fixedValue: date, modified: modified };
		}
	}
};

/**
 * Iterates over all Google Sheets in a folder and executes a callback for each.
 */
function forEachSpreadsheetInFolder(callback)
{
	const folder = DriveApp.getFolderById(FOLDER_ID);
	const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

	while (files.hasNext())
	{
		const file = files.next();
		const ssName = file.getName();
		
		try
		{
			const ss = SpreadsheetApp.open(file);
			const ssUrl = ss.getUrl();
			callback(ss, ssName, ssUrl);
		}
		catch (e)
		{
			console.error(`❌ Erreur critique sur "${ssName}" : ${e.toString()}`);
		}
	}
}

function dryRunFixClericalErrors()
{
	processClericalErrors(true);
}

function applyFixClericalErrors()
{
	const ui = SpreadsheetApp.getUi();
	const response = ui.alert(
		'Confirmation',
		'Êtes-vous sûr de vouloir appliquer les corrections sur TOUS les fichiers ?',
		ui.ButtonSet.YES_NO
	);

	if (response === ui.Button.YES)
	{
		processClericalErrors(false);
	}
}

/**
 * Iterates through files and applies the corrections defined in FIX_CONFIG.
 * @param {boolean} isDryRun If true, does not modify files.
 */
function processClericalErrors(isDryRun)
{
	const modeLabel = isDryRun ? '[DRY RUN]' : '[LIVE]';
	const logData = [['Horodatage', 'Niveau', 'Fichier', 'Message']];
	
	const log = (message, level = 'INFO', fileName = '-') =>
	{
		const timestamp = Utilities.formatDate(new Date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
		logData.push([timestamp, level, fileName, message]);
		if (level === 'ERROR')
		{
			console.error(`${modeLabel} ${fileName !== '-' ? '[' + fileName + '] ' : ''}${message}`);
		}
		else
		{
			console.info(`${modeLabel} ${fileName !== '-' ? '[' + fileName + '] ' : ''}${message}`);
		}
	};

	log(`Début du traitement des erreurs cléricales.`);

	let filesProcessed = 0;
	let errorsCorrected = 0;
	let errorsUnfixable = 0;

	forEachSpreadsheetInFolder((ss, ssName, ssUrl) => 
	{
		log(`Traitement du fichier (${ssUrl})`, 'INFO', ssName);
		filesProcessed++;
		let fileHasChanges = false;

		for (const address in FIX_CONFIG)
		{
			const config = FIX_CONFIG[address];
			const range = ss.getRange(address);
			
			if (!range)
			{
				log(`${config.name} : ❌ Impossible d'accéder à "${address}"`, 'ERROR', ssName);
				errorsUnfixable++;
				continue;
			}

			const oldValue = range.getValue();
			const result = config.fixer(oldValue, ss);

			if (result.error)
			{
				log(`${config.name} : ❌ ERREUR NON CORRIGÉE : ${result.error}`, 'ERROR', ssName);
				errorsUnfixable++;
			}
			else if (result.modified)
			{
				const actionLabel = isDryRun ? 'SIMULATION' : 'CORRECTION';
				const oldDisplayValue = (oldValue instanceof Date) ? Utilities.formatDate(oldValue, Session.getScriptTimeZone(), 'dd/MM/yyyy') : oldValue;
				const newDisplayValue = (result.fixedValue instanceof Date) ? Utilities.formatDate(result.fixedValue, Session.getScriptTimeZone(), 'dd/MM/yyyy') : result.fixedValue;
				
				log(`${config.name} : ✅ ${actionLabel} de "${oldDisplayValue}" vers "${newDisplayValue}"`, 'INFO', ssName);
				if (!isDryRun)
				{
					range.setValue(result.fixedValue);
					fileHasChanges = true;
				}
				errorsCorrected++;
			}
		}

		if (fileHasChanges && !isDryRun)
		{
			SpreadsheetApp.flush();
		}
	});

	const totalErrorsFound = errorsCorrected + errorsUnfixable;
	const summary = `Traitement terminé.
Fichiers parcourus : ${filesProcessed}
Erreurs trouvées : ${totalErrorsFound}
Erreurs corrigées : ${errorsCorrected}
Erreurs restantes : ${errorsUnfixable}`;

	log(summary);
	createLogSheet(logData, isDryRun, summary);
	SpreadsheetApp.getUi().alert(`${modeLabel} ${summary}`);
}

/**
 * Creates a log sheet in the active spreadsheet.
 * @param {Array<Array<string>>} data
 * @param {boolean} isDryRun
 * @param {string} summary
 */
function createLogSheet(data, isDryRun, summary)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const timestamp = Utilities.formatDate(new Date, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
	const mode = isDryRun ? 'DRY' : 'LIVE';
	const sheetName = `Logs_${mode}_${timestamp}`;
	
	const sheet = ss.insertSheet(sheetName);
	
	// Prepare the summary at the top
	const summaryRows = summary.split('\n').map(line =>
	{
		return [line];
	});
	sheet.getRange(1, 1, summaryRows.length, 1).setValues(summaryRows).setFontWeight('bold');
	
	// Inject log data after the summary (+ 1 empty line)
	const startRow = summaryRows.length + 2;
	const range = sheet.getRange(startRow, 1, data.length, data[0].length);
	range.setValues(data);

	// Format the table header
	sheet.getRange(startRow, 1, 1, data[0].length).setFontWeight('bold').setBackground('#f3f3f3');
	sheet.setFrozenRows(startRow);
	sheet.autoResizeColumns(1, data[0].length);

	// Reduce sheet size to match data
	const totalRowsUsed = startRow + data.length - 1;
	if (sheet.getMaxRows() > totalRowsUsed)
	{
		sheet.deleteRows(totalRowsUsed + 1, sheet.getMaxRows() - totalRowsUsed);
	}
	if (sheet.getMaxColumns() > data[0].length)
	{
		sheet.deleteColumns(data[0].length + 1, sheet.getMaxColumns() - data[0].length);
	}
}

/**
 * Collects structures data from individual files into the 'Structures' sheet.
 * Normalizes dates to 9:00 AM (.375 decimal part).
 */
function collectStructuresFromIndividualFiles()
{
	const activeSs = SpreadsheetApp.getActiveSpreadsheet();
	const targetSheet = activeSs.getSheetByName('Structures');

	if (!targetSheet)
	{
		console.error("❌ Erreur : La feuille 'Structures' n'existe pas.");
		return;
	}

	const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
	console.info("ℹ️ Début du rechargement des données.");

	forEachSpreadsheetInFolder((ss, ssName, ssUrl) => 
	{
		console.info(`ℹ️ Traitement du fichier : ${ssName} (${ssUrl})`);
		const sourceSheet = ss.getSheetByName('Données');

		if (!sourceSheet)
		{
			console.warn(`⚠️ Pas de feuille 'Données' dans ${ssName}`);
			return;
		}

		const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

		if (JSON.stringify(sourceHeaders) !== JSON.stringify(targetHeaders))
		{
			console.error(`❌ En-têtes invalides dans ${ssName}`);
			return;
		}

		const sourceDataRow = sourceSheet.getRange(2, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
		const isEmpty = sourceDataRow.every(cell =>
		{
			return cell === '' || cell === null;
		});
		
		if (isEmpty)
		{
			console.error(`❌ Ligne 2 vide dans ${ssName}`);
			return;
		}

		// Normalize dates to 9:00 AM
		const normalizedRow = sourceDataRow.map(normalizeDate);

		targetSheet.appendRow(normalizedRow);
		console.log(`✅ Données importées : ${ssName}`);
	});
}

/**
 * Collects personnes data from individual files into the 'Personnes' sheet.
 * Normalizes dates to 9:00 AM (.375 decimal part).
 */
function collectPersonnesFromIndividualFiles()
{
	const activeSs = SpreadsheetApp.getActiveSpreadsheet();
	const targetSheet = activeSs.getSheetByName('Personnes');

	if (!targetSheet)
	{
		console.error("❌ Erreur : La feuille 'Personnes' n'existe pas.");
		return;
	}

	console.info("ℹ️ Début de l'importation des données Personnes.");

	let fileCount = 0;
	forEachSpreadsheetInFolder((ss, ssName, ssUrl) =>
	{
		fileCount++;
		console.info(`ℹ️ [${fileCount}] Traitement du fichier : ${ssName} (${ssUrl})`);

		const saisieSheet = ss.getSheetByName('Saisie');
		if (!saisieSheet)
		{
			console.warn(`⚠️ Pas de feuille 'Saisie' dans ${ssName}`);
			return;
		}

		const codeBA = saisieSheet.getRange('C2').getValue();
		const nomPartenaire = saisieSheet.getRange('C3').getValue();
		console.log(`DEBUG: Code BA="${codeBA}", Partenaire="${nomPartenaire}"`);

		// Reading fixed range B11:F26
		const dataRange = saisieSheet.getRange('B11:F26');
		const values = dataRange.getValues();
		console.log(`DEBUG: Plage B11:F26 lue, nombre de lignes: ${values.length}`);

		let rowsFoundInFile = 0;
		for (let i = 0; i < values.length; i++)
		{
			const row = values[i];
			const prenomNom = row[1]; // Index 1 corresponds to column C in range B:F

			if (prenomNom && String(prenomNom).trim() !== '')
			{
				const rowToAppend = [codeBA, nomPartenaire].concat(row);
				
				// Normalize dates to 9:00 AM
				const normalizedRowToAppend = rowToAppend.map(normalizeDate);

				console.log(`DEBUG: Tentative d'ajout pour "${prenomNom}" : ${JSON.stringify(normalizedRowToAppend)}`);
				
				try
				{
					targetSheet.appendRow(normalizedRowToAppend);
					const lastRow = targetSheet.getLastRow();
					console.info(`✅ Personne importée depuis ${ssName} : ${prenomNom} (Ajoutée à la ligne ${lastRow})`);
					rowsFoundInFile++;
				}
				catch (e)
				{
					console.error(`❌ Erreur lors de l'ajout de la ligne pour "${prenomNom}" dans ${ssName} : ${e.toString()}`);
				}
			}
		}

		console.log(`DEBUG: Fichier "${ssName}" terminé. Lignes ajoutées: ${rowsFoundInFile}`);
	});

	console.info(`ℹ️ Fin de l'importation. Total de fichiers traités: ${fileCount}`);
}
