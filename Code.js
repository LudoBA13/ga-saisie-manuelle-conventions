const FOLDER_ID = '1WoTjgUiF4PjVQNBX6chRQzDFUP2f7k3a';

function onOpen()
{
	SpreadsheetApp.getUi()
		.createMenu('Conventions')
		.addItem('Recharger les données des conventions', 'rechargerDonneesConventions')
		.addSeparator()
		.addItem('Simuler la correction des erreurs (Dry Run)', 'dryRunFixClericalErrors')
		.addItem('Appliquer la correction des erreurs', 'applyFixClericalErrors')
		.addToUi();
}

/**
 * Configuration des corrections automatiques.
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

			if (/^\d+$/.test(strValue))
			{
				return { success: true, fixedValue: strValue, modified: false };
			}

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

			let modified = false;
			const originalInput = value;

			// Tentative de correction si le format est DDMMYY ou DDMMYYYY (sans les slashs)
			// La regex demandée par l'utilisateur : ^([0-9]{1,2})([0-9]{1,2})([0-9]{4}|[0-9]{2})$
			const dateRegex = /^([0-9]{1,2})([0-9]{1,2})([0-9]{4}|[0-9]{2})$/;
			const match = strValue.match(dateRegex);
			if (match)
			{
				strValue = `${match[1]}/${match[2]}/${match[3]}`;
				modified = true;
			}

			// Analyse de la date
			// Note: Google Apps Script / JS Date peut être capricieux avec le format DD/MM/YYYY.
			// On va parser manuellement pour être sûr.
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
			const now = new Date();
			const minDate = new Date(2024, 0, 1);

			// Vérification de la validité réelle de la date (ex: pas de 31 février)
			if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day)
			{
				return { error: `La date "${strValue}" est calendairement invalide.` };
			}

			if (date < minDate || date > now)
			{
				return { error: `La date "${strValue}" est hors limites (doit être entre 01/01/2024 et aujourd'hui).` };
			}

			const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
			
			// Si la valeur a été reformattée ou corrigée par regex
			if (modified || formattedDate !== originalValue)
			{
				return { success: true, fixedValue: formattedDate, modified: true };
			}

			return { success: true, fixedValue: formattedDate, modified: false };
		}
	}
};

/**
 * Itère sur tous les Google Sheets d'un dossier et exécute un callback pour chacun.
 * @param {function(GoogleAppsScript.Spreadsheet.Spreadsheet, string)} callback 
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
			callback(ss, ssName);
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

function processClericalErrors(isDryRun)
{
	const modeLabel = isDryRun ? '[DRY RUN]' : '[LIVE]';
	console.info(`${modeLabel} Début du traitement des erreurs cléricales.`);

	let filesProcessed = 0;
	let errorsCorrected = 0;
	let errorsUnfixable = 0;

	forEachSpreadsheetInFolder((ss, ssName) => 
	{
		console.info(`${modeLabel} ℹ️ Traitement du fichier : ${ssName}`);
		filesProcessed++;
		let fileHasChanges = false;

		for (const address in FIX_CONFIG)
		{
			const config = FIX_CONFIG[address];
			const range = ss.getRange(address);
			
			if (!range)
			{
				console.error(`${modeLabel} [${ssName}] ❌ Impossible d'accéder à "${address}"`);
				errorsUnfixable++;
				continue;
			}

			const oldValue = range.getValue();
			const result = config.fixer(oldValue);

			if (result.error)
			{
				console.error(`${modeLabel} [${ssName}] ${config.name} : ❌ ERREUR NON CORRIGÉE : ${result.error}`);
				errorsUnfixable++;
			}
			else if (result.modified)
			{
				const actionLabel = isDryRun ? 'SIMULATION' : 'CORRECTION';
				console.info(`${modeLabel} [${ssName}] ${config.name} : ✅ ${actionLabel} de "${oldValue}" vers "${result.fixedValue}"`);
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

	console.info(`${modeLabel} ${summary}`);
	SpreadsheetApp.getUi().alert(`${modeLabel} ${summary}`);
}

function rechargerDonneesConventions()
{
	const activeSs = SpreadsheetApp.getActiveSpreadsheet();
	const targetSheet = activeSs.getSheetByName('Data');

	if (!targetSheet)
	{
		console.error("❌ Erreur : La feuille 'Data' n'existe pas.");
		return;
	}

	const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
	console.info("ℹ️ Début du rechargement des données.");

	forEachSpreadsheetInFolder((ss, ssName) => 
	{
		console.info(`ℹ️ Traitement du fichier : ${ssName}`);
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
		const isEmpty = sourceDataRow.every(cell => cell === '' || cell === null);
		
		if (isEmpty)
		{
			console.error(`❌ Ligne 2 vide dans ${ssName}`);
			return;
		}

		targetSheet.appendRow(sourceDataRow);
		console.log(`✅ Données importées : ${ssName}`);
	});
}
