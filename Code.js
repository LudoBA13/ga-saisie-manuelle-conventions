const FOLDER_ID = '1WoTjgUiF4PjVQNBX6chRQzDFUP2f7k3a';

function onOpen()
{
	SpreadsheetApp.getUi()
		.createMenu('Conventions')
		.addItem('Importer données structures', 'rechargerDonneesConventions')
		.addSeparator()
		.addItem('Simuler la correction des erreurs (Dry Run)', 'dryRunFixClericalErrors')
		.addItem('Appliquer la correction des erreurs', 'applyFixClericalErrors')
		.addToUi();
}

/**
 * Configuration des corrections automatiques.
 * Chaque clé est une adresse de cellule (Feuille!A1).
 * Chaque valeur est un objet avec un nom et une fonction 'fixer'.
 */
const FIX_CONFIG = {
	'Saisie!C2': {
		name: 'Numéro de convention',
		fixer: (value) =>
		{
			const strValue = String(value || '').trim();
			if (strValue === '')
			{
				return { error: 'La valeur est vide.' };
			}

			// Si c'est déjà un entier pur
			if (/^\d+$/.test(strValue))
			{
				return { success: true, fixedValue: strValue, modified: false };
			}

			// Application de la regex RE2 : ^[Bb]*0*([1-9][0-9]+).* -> $1
			const regex = /^[Bb]*0*([1-9][0-9]+).*/;
			const match = strValue.match(regex);

			if (match && match[1])
			{
				const fixed = match[1];
				if (/^\d+$/.test(fixed))
				{
					return { success: true, fixedValue: fixed, modified: true };
				}
			}

			return { error: `La valeur "${strValue}" ne correspond pas à un format d'entier valide.` };
		}
	},
	'Saisie!C72': {
		name: 'Date de signature',
		fixer: (value) =>
		{
			let strValue;
			if (value instanceof Date)
			{
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
			let modified = false;

			// Tentative de correction si le format est D/MYYYY ou DD/MMYYYY (manque le deuxième slash)
			const dateRegex = /^([0-9]{1,2})\/([0-9]{1,2})([0-9]{4}|[0-9]{2})$/;
			const match = strValue.match(dateRegex);
			if (match)
			{
				strValue = `${match[1]}/${match[2]}/${match[3]}`;
				modified = true;
			}

			// Analyse de la date
			// Note: Google Apps Script / JS Date peut être capricieux avec le format DD/MM/YYYY.
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

			const date = new Date(year, month, day);
			const now = new Date;
			const minDate = new Date(2024, 0, 1);

			// Vérification de la validité réelle de la date
			if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day)
			{
				return { error: `La date "${strValue}" est calendairement invalide.` };
			}

			if (date < minDate || date > now)
			{
				return { error: `La date "${strValue}" est hors limites (doit être entre 01/01/2024 et aujourd'hui).` };
			}

			const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
			
			// Si la valeur a été reformattée ou corrigée
			if (modified || formattedDate !== originalStrValue)
			{
				return { success: true, fixedValue: formattedDate, modified: true };
			}

			return { success: true, fixedValue: formattedDate, modified: false };
		}
	}
};

/**
 * Itère sur tous les Google Sheets d'un dossier et exécute un callback pour chacun.
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
 * Parcourt les fichiers et applique les corrections définies dans FIX_CONFIG.
 * @param {boolean} isDryRun Si vrai, ne modifie pas les fichiers.
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
			const result = config.fixer(oldValue);

			if (result.error)
			{
				log(`${config.name} : ❌ ERREUR NON CORRIGÉE : ${result.error}`, 'ERROR', ssName);
				errorsUnfixable++;
			}
			else if (result.modified)
			{
				const actionLabel = isDryRun ? 'SIMULATION' : 'CORRECTION';
				log(`${config.name} : ✅ ${actionLabel} de "${oldValue}" vers "${result.fixedValue}"`, 'INFO', ssName);
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
 * Crée une feuille de logs dans le classeur actif.
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
	
	// Préparation du résumé en haut
	const summaryRows = summary.split('\n').map(line =>
	{
		return [line];
	});
	sheet.getRange(1, 1, summaryRows.length, 1).setValues(summaryRows).setFontWeight('bold');
	
	// Injection des données de logs après le résumé (+ 1 ligne vide)
	const startRow = summaryRows.length + 2;
	const range = sheet.getRange(startRow, 1, data.length, data[0].length);
	range.setValues(data);

	// Mise en forme de l'en-tête du tableau
	sheet.getRange(startRow, 1, 1, data[0].length).setFontWeight('bold').setBackground('#f3f3f3');
	sheet.setFrozenRows(startRow);
	sheet.autoResizeColumns(1, data[0].length);

	// Réduction de la taille de la feuille pour correspondre aux données
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
 * Recharge les données des conventions dans la feuille 'Data'.
 */
function rechargerDonneesConventions()
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

		targetSheet.appendRow(sourceDataRow);
		console.log(`✅ Données importées : ${ssName}`);
	});
}
