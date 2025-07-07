var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();


//the main method, calls the two important methods

function main() {

hoursPerDay();

totalHoursPerPerson();

}

// Trigger function to run hoursPerDay when new form response is submitted
// Also, in case you are curious, yes this code is unneeded. The trigger could just be set to hoursPerDay(), and it would run just the same.
// This is just an extra call for no reason, but who cares about optimization anyways? This is JS, not C.
// Future me here, I have no idea what past me meant. hoursPerDay doesn't call totalHoursPerPerson, and they both need called for every shift
// change, so I have no idea what he really meant.
// This way of calling actually saves nanosecond level performance over calling main directly. If we called main() just to call hoursPerDay() and totatlHoursPerPerson(), then JS engine would allocate a stack frame for main, as well as one for hoursPerDay and totalHoursPerPerson. By just calling them directly, we save a creation and deletion of a stack frame, saving some time and some memory. However, if main were frequently called in other places, the JS engine might inline it, replacing the call to main() with its contents directly in the calling function. This optimization would eliminate the need for a separate stack frame, and the performance difference would disappear. But, since main doesn't really do anything in this program, this way is technically faster, even if its only by about 10-20 nanoseconds.

function onFormSubmit(e) {

hoursPerDay(); // Run the function every time a new response is submitted

totalHoursPerPerson(); // Run the function to calculate # of hours everytime a new response is submitted

}


function hoursPerDay() {

var startColumn = 4; 

var lastRow = sheet.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();


for (var col = startColumn; col <= sheet.getLastColumn(); col++) {

var totalHours = 0; // Accumulator for total hours for the day

Logger.log(`\nProcessing Column: ${col}`);


for (var rowNumber = 2; rowNumber <= lastRow; rowNumber++) { // Start at row 2

var cellValue = sheet.getRange(rowNumber, col).getValue().toString().trim();


if (cellValue === "") {

Logger.log(`Row ${rowNumber}, Column ${col}: Skipping empty cell`);

continue; // Skip empty rows

}


try {

// Extract valid shifts from the cell

const validShifts = extractValidShifts(cellValue);

Logger.log(`Row ${rowNumber}, Column ${col}: Extracted Valid Shifts - "${validShifts}"`);


// Compact multiple shifts into one

const compactedShift = compactShifts(validShifts);

Logger.log(`Row ${rowNumber}, Column ${col}: Compact Shift Result - "${compactedShift}"`);


// Validate compactedShift before calling convertTime

if (!compactedShift || compactedShift.trim() === "") {

Logger.log(`Row ${rowNumber}, Column ${col}: Invalid compacted shift: "${compactedShift}"`);

continue; // Skip invalid shifts

}


// Calculate hours worked

const hours = convertTime(compactedShift);

Logger.log(`Row ${rowNumber}, Column ${col}: Hours from Shift - "${hours}"`);


totalHours += hours; // Add hours for the individual to total

} catch (error) {

Logger.log(`Row ${rowNumber}, Column ${col}: Error - ${error.message}`);

}

}


// Log total hours for debugging

Logger.log(`Column ${col}: Total Hours for Day - ${totalHours}`);


// Write the total hours for the day in the last row of the column

sheet.getRange(lastRow + 1, col).setValue("Total hours for this day is: " + Math.round(totalHours));

}

}


function totalHoursPerPerson() {

var startColumn = 4; // Column D is the first relevant column (after column C)

var lastRow = sheet.getRange(1, 1).getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();

var lastColumn = sheet.getLastColumn();

var totalHours = {};


// Check if "Total hours" column exists

var headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

var totalHoursColIndex = headerRow.indexOf("Total hours:") + 1; // Add 1 because indexOf is 0-based


// If "Total hours" column does not exist, create it

if (totalHoursColIndex === 0) {

totalHoursColIndex = lastColumn + 1;

sheet.getRange(1, totalHoursColIndex).setValue("Total hours:");

}


for (var col = startColumn; col <= lastColumn; col++) {

var headerValue = sheet.getRange(1, col).getValue().toString().toLowerCase();

if (headerValue.startsWith("available")) {

break;

}


for (var rowNumber = 2; rowNumber <= lastRow; rowNumber++) {

var personName = sheet.getRange(rowNumber, 3).getValue(); // Names are in column C

var cellValue = sheet.getRange(rowNumber, col).getValue().toString().trim();


if (!totalHours[personName]) {

totalHours[personName] = 0;

}


if (cellValue.match(/am|pm/) && cellValue.includes(":") && cellValue.includes("-")) {

try {

// Extract valid shifts from the cell

const validShifts = extractValidShifts(cellValue);

Logger.log(`Row ${rowNumber}, Column ${col}: Extracted Valid Shifts - "${validShifts}"`);


// Compact multiple shifts into one

const compactedShift = compactShifts(validShifts);

Logger.log(`Row ${rowNumber}, Column ${col}: Compact Shift Result - "${compactedShift}"`);


// Validate compactedShift before calling convertTime

if (!compactedShift || compactedShift.trim() === "") {

Logger.log(`Row ${rowNumber}, Column ${col}: Invalid compacted shift: "${compactedShift}"`);

continue; // Skip invalid shifts

}


// Calculate hours worked

const hours = convertTime(compactedShift);

Logger.log(`Row ${rowNumber}, Column ${col}: Hours from Shift - "${hours}"`);


totalHours[personName] += hours;

} catch (error) {

Logger.log(`Row ${rowNumber}, Column ${col}: Error - ${error.message}`);

}

}

}

}


// Write total hours to the "Total hours:" column in the same row as the person

for (var rowNumber = 2; rowNumber <= lastRow; rowNumber++) {

var personName = sheet.getRange(rowNumber, 3).getValue(); // Names are in column C

if (totalHours[personName] !== undefined) {

sheet.getRange(rowNumber, totalHoursColIndex).setValue(Math.round(totalHours[personName]));

Logger.log(`Total hours for "${personName}" set to ${Math.round(totalHours[personName])} hours in row ${rowNumber}`);

}

}

}


// Converts time from 12 hours to 24 hours for easier hours worked calculation

function convertTime(shift) {

// Validate shift format first. Checks for the presence of am/pm, colon (:), dash (-), and ensures the string

// starts and ends with valid times in the format HH:MM (hours are 1-12, minutes are 00-59) followed by am or pm.

if (!shift.match(/^\d{1,2}:\d{2}(am|pm)-\d{1,2}:\d{2}(am|pm)$/)) {

throw new Error("Invalid shift formatting");

}


// we want time in the 24 hour clock, but are getting shifts in the 12 hour format, 

// so, we need to convert to 24 hour time.


// break the shift down into two parts for easier conversion

var shiftParts = shift.split("-")



// if the shift takes place in the afternoon (pm), need to convert it to 24 hour time

// run for both shift parts

for(var i = 0; i < 2; i++) { 

// if its 12am, make it 00 (it won't ever be 12am for day shift, but if night shift gets this code somehow, I put in an edge case for them)

if (shiftParts[i].match("12") && shiftParts[i].match("am")) {

shiftParts[i] = shiftParts[i].replace("12", "0");

}

if(shiftParts[i].match("pm")) {

// if its 12pm, don't add anything since the number is already correct in 24 hour time

if (shiftParts[i].match("12")) { 

}

// else add 12 hours for time between 1pm and 11:59pm.

else {

// make a variable that stores the hour of the shift by removing everything after the colon using substring

var tempShift = shiftParts[i].substring(0, shiftParts[i].indexOf(":"));

// get it as an int for addition

var shiftHour = parseInt(tempShift);

// add 12 hours to it

shiftHour = shiftHour + 12;

// put the shifts hour back into the shift. tempShift is the variable that holds the hour of the shift, so we just replace that with our new shift time as a string

shiftParts[i] = shiftParts[i].replace(tempShift, shiftHour.toString());

}


// its am, don't do anything to the hours

}


// now, remove the shift parts am/pm labels

shiftParts[i] = shiftParts[i].replace("pm", "");

shiftParts[i] = shiftParts[i].replace("am", ""); 

}

// the hours should now be in 24 hour time, so convert them to integers 


// Split the time string into hours and minutes, where the startShift and endShift variables represent the time as a decimal (12.5 vs 12:30pm)

var [hours, minutes] = shiftParts[0].split(":").map(Number);


// Convert the time to a decimal value

var startShift = hours + (minutes / 60);


// repeat for the other shift

[hours, minutes] = shiftParts[1].split(":").map(Number);

var endShift = hours + (minutes / 60);

// now subtract them to get the number of hours worked

var hoursWorked = Math.abs(endShift - startShift);


// since there is a mandatory lunch from 11:30am to 12:00pm, we need to remove 0.5 hours from every shift that crosses that time frame.

// as a decimal, 11:30am is 11.5 and 12:00pm is 12, so any shift that starts before 11.5 and ends after 12 has that lunch break, but no other shifts. Ex: 10am (10) to 2pm (14) has a lunch, but 7am (7) to 11am (11), does not.

if(startShift <= 11.5 && endShift >= 12) {

// using precise rounding for completeness sake, even though everyone works in 15 minute increments

hoursWorked = Math.round((hoursWorked - 0.5) * 100) / 100;

}

// return the hours worked so the program actually functions

return hoursWorked;

}




// Helper function to make shifts go from long lists to one "compact" shift

function compactShifts(shiftList) {

// Split the list into individual shifts and trim whitespace

var shifts = shiftList.split(",").map(shift => shift.trim());


var earliestStart = null;

var latestEnd = null;


// Iterate through all shifts to find the earliest start and latest end times

shifts.forEach(function (shift) {

if (!shift.match(/^\d{1,2}:\d{2}(am|pm)-\d{1,2}:\d{2}(am|pm)$/)) {

throw new Error(`Invalid shift format: "${shift}"`);

}


var [start, end] = shift.split("-");


// Compare and set the earliest start time

if (!earliestStart || isEarlier(start, earliestStart)) {

earliestStart = start;

}


// Compare and set the latest end time

if (!latestEnd || isEarlier(latestEnd, end)) {

latestEnd = end;

}

});


// Return the compacted shift

return `${earliestStart}-${latestEnd}`;

}


// Helper function to compare two times in HH:MMam/pm format

function isEarlier(time1, time2) {

// Convert times to 24-hour format using a simplified version of convertTime logic

var [hour1, minute1] = time1.match(/(\d{1,2}):(\d{2})(am|pm)/).slice(1, 4);

var [hour2, minute2] = time2.match(/(\d{1,2}):(\d{2})(am|pm)/).slice(1, 4);


// Convert to integers

hour1 = parseInt(hour1, 10);

hour2 = parseInt(hour2, 10);

minute1 = parseInt(minute1, 10);

minute2 = parseInt(minute2, 10);


// Handle AM/PM conversions

if (time1.includes("pm") && hour1 !== 12) hour1 += 12;

if (time2.includes("pm") && hour2 !== 12) hour2 += 12;

if (time1.includes("am") && hour1 === 12) hour1 = 0;

if (time2.includes("am") && hour2 === 12) hour2 = 0;


// Compare times

if (hour1 < hour2 || (hour1 === hour2 && minute1 < minute2)) {

return true;

}

return false;

}


// Helper function to look through the "other" responses and see if they have valid shift data (eg "I can do 7:30am-10:30am" has the valid shift

// "7:30am-10:30am"). If it does, return it.

function extractValidShifts(cellValue) {

// Regex to match valid shifts in the format HH:MM(am|pm)-HH:MM(am|pm)

const shiftRegex = /\b\d{1,2}:\d{2}(am|pm)-\d{1,2}:\d{2}(am|pm)\b/g;


// Find all valid shifts in the cell value

const matches = cellValue.match(shiftRegex);


if (!matches || matches.length === 0) {

throw new Error(`No valid shifts found in cell value: "${cellValue}"`);

}


// Join valid shifts with ", " for consistent formatting

return matches.join(", ");

}


