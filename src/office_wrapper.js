export class ExcelWrapper {
  static officeIsReady = false;

    // Check if Office.js is available and we're running in an Office context
    static initializeOffice() {
        console.log("Checking Office.js availability...");

        if (typeof Office === 'undefined' || typeof Office.onReady !== "function") {
            console.log("Office.js library not loaded");
            return false;
        }

        void Office.onReady(() => {
            console.log("Office.js library loaded");
            ExcelWrapper.officeIsReady = true;
        });

        return ExcelWrapper.officeIsReady;
    }


    // Get the value from a specific cell address (e.g., "A1", "B2")
    static async getCellValue(address) {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const cell = sheet.getRange(address);
                cell.load("values");
                await context.sync();

                const value = cell.values[0][0];
                return value !== null && value !== undefined ? String(value) : "";
            });
        } catch (error) {
            throw new Error(`Failed to get cell value from ${address}: ${error.message}`);
        }
    }

    // Set the value of a specific cell address
    static async setCellValue(address, value) {
        try {
            await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const cell = sheet.getRange(address);
                cell.values = [[value]];
                await context.sync();
            });
            return true;
        } catch (error) {
            throw new Error(`Failed to set cell value at ${address}: ${error.message}`);
        }
    }

    // Get the name of the active worksheet
    static async getActiveSheetName() {
        try {
            return await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.load("name");
                await context.sync();
                return sheet.name;
            });
        } catch (error) {
            throw new Error(`Failed to get sheet name: ${error.message}`);
        }
    }
}

// Make it available globally as well for easier debugging
window.ExcelWrapper = ExcelWrapper;
