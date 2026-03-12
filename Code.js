const FOLDER_ID = '1WoTjgUiF4PjVQNBX6chRQzDFUP2f7k3a';

function onOpen()
{
	SpreadsheetApp.getUi()
		.createMenu('Conventions')
		.addItem('Reload convention data', 'rechargerDonneesConventions')
		.addSeparator()
		.addItem('Simulate clerical error fix (Dry Run)', 'dryRunFixClericalErrors')
		.addItem('Apply clerical error fix', 'applyFixClericalErrors')
		.addToUi();
}

/**
 * Configuration for automatic fixes.
 * Each key is a cell address (Sheet!A1).
 * Each value is an object with a name and a 'fixer' function.
 */
const FIX_CONFIG = {
	'Saisie!C2': {
		name: 'Convention Number',
		fixer: (value) =>
		{
			const strValue = String(value || '').trim();
			if (strValue === '')
			{
				return { error: 'Value is empty.' };
			}

			// If it's already a pure integer
			if (/^\d+$/.test(strValue))
			{
				return { success: true, fixedValue: strValue, modified: false };
			}

			// Apply RE2 regex: ^[Bb]*0*([1-9][0-9]+).* -> $1
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

			return { error: `Value "${strValue}" does not match a valid integer format.` };
		}
	},
	'Saisie!C72': {
		name: 'Signature Date',
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
				return { error: 'Value is empty.' };
			}

			const originalStrValue = strValue;
			let modified = false;

			// Attempt correction if format is D/MYYYY or DD/MMYYYY (missing the second slash)
			const dateRegex = /^([0-9]{1,2})\/([0-9]{1,2})([0-9]{4}|[0-9]{2})$/;
			const match = strValue.match(dateRegex);
			if (match)
			{
				strValue = `${match[1]}/${match[2]}/${match[3]}`;
				modified = true;
			}

			// Parse the date
			// Note: JS Date can be unreliable with DD/MM/YYYY format, so we parse manually.
			const parts = strValue.split(/[\/\-\.]/);
			if (parts.length !== 3)
			{
				return { error: `Invalid date format: "${strValue}". Expected: DD/MM/YYYY.` };
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

			// Verify if the date is actually valid (e.g., no Feb 31st)
			if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day)
			{
				return { error: `Date "${strValue}" is calendrically invalid.` };
			}

			if (date < minDate || date > now)
			{
				return { error: `Date "${strValue}" is out of bounds (must be between 01/01/2024 and today).` };
			}

			const formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
			
			// If the value was reformatted or fixed via regex
			if (modified || formattedDate !== originalStrValue)
			{
				return { success: true, fixedValue: formattedDate, modified: true };
			}

			return { success: true, fixedValue: formattedDate, modified: false };
		}
	}
};

/**
 * Iterates through all Google Sheets in a folder and executes a callback for each.
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
			console.error(`❌ Critical error on "${ssName}": ${e.toString()}`);
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
		'Are you sure you want to apply fixes to ALL files?',
		ui.ButtonSet.YES_NO
	);

	if (response === ui.Button.YES)
	{
		processClericalErrors(false);
	}
}

/**
 * Processes clerical errors across all spreadsheets based on FIX_CONFIG.
 * @param {boolean} isDryRun If true, does not modify the files.
 */
function processClericalErrors(isDryRun)
{
	const modeLabel = isDryRun ? '[DRY RUN]' : '[LIVE]';
	console.info(`${modeLabel} Starting clerical error processing.`);

	let filesProcessed = 0;
	let errorsCorrected = 0;
	let errorsUnfixable = 0;

	forEachSpreadsheetInFolder((ss, ssName) => 
	{
		console.info(`${modeLabel} ℹ️ Processing file: ${ssName}`);
		filesProcessed++;
		let fileHasChanges = false;

		for (const address in FIX_CONFIG)
		{
			const config = FIX_CONFIG[address];
			const range = ss.getRange(address);
			
			if (!range)
			{
				console.error(`${modeLabel} [${ssName}] ❌ Cannot access "${address}"`);
				errorsUnfixable++;
				continue;
			}

			const oldValue = range.getValue();
			const result = config.fixer(oldValue);

			if (result.error)
			{
				console.error(`${modeLabel} [${ssName}] ${config.name}: ❌ ERROR NOT FIXED: ${result.error}`);
				errorsUnfixable++;
			}
			else if (result.modified)
			{
				const actionLabel = isDryRun ? 'SIMULATION' : 'FIXED';
				console.info(`${modeLabel} [${ssName}] ${config.name}: ✅ ${actionLabel} from "${oldValue}" to "${result.fixedValue}"`);
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
	const summary = `Processing completed.
Files scanned: ${filesProcessed}
Errors found: ${totalErrorsFound}
Errors fixed: ${errorsCorrected}
Errors remaining: ${errorsUnfixable}`;

	console.info(`${modeLabel} ${summary}`);
	SpreadsheetApp.getUi().alert(`${modeLabel} ${summary}`);
}

/**
 * Reloads convention data into the 'Data' sheet.
 */
function rechargerDonneesConventions()
{
	const activeSs = SpreadsheetApp.getActiveSpreadsheet();
	const targetSheet = activeSs.getSheetByName('Data');

	if (!targetSheet)
	{
		console.error("❌ Error: 'Data' sheet does not exist.");
		return;
	}

	const targetHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
	console.info("ℹ️ Starting data reload.");

	forEachSpreadsheetInFolder((ss, ssName) => 
	{
		console.info(`ℹ️ Processing file: ${ssName}`);
		const sourceSheet = ss.getSheetByName('Données');

		if (!sourceSheet)
		{
			console.warn(`⚠️ No 'Données' sheet in ${ssName}`);
			return;
		}

		const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];

		if (JSON.stringify(sourceHeaders) !== JSON.stringify(targetHeaders))
		{
			console.error(`❌ Invalid headers in ${ssName}`);
			return;
		}

		const sourceDataRow = sourceSheet.getRange(2, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
		const isEmpty = sourceDataRow.every(cell => cell === '' || cell === null);
		
		if (isEmpty)
		{
			console.error(`❌ Row 2 is empty in ${ssName}`);
			return;
		}

		targetSheet.appendRow(sourceDataRow);
		console.log(`✅ Data imported: ${ssName}`);
	});
}
