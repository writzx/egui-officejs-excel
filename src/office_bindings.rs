use wasm_bindgen::prelude::*;
use wasm_bindgen_futures::JsFuture;
use web_sys::console;

// Import our JavaScript wrapper functions
#[wasm_bindgen(module = "/src/office_wrapper.js")]
extern "C" {
    #[wasm_bindgen(js_namespace = ExcelWrapper)]
    fn initializeOffice() -> bool;

    #[wasm_bindgen(js_namespace = ExcelWrapper)]
    fn getCellValue(address: &str) -> js_sys::Promise;

    #[wasm_bindgen(js_namespace = ExcelWrapper)]
    fn setCellValue(address: &str, value: &str) -> js_sys::Promise;

    #[wasm_bindgen(js_namespace = ExcelWrapper)]
    fn getActiveSheetName() -> js_sys::Promise;
}

// Rust wrapper for Office.js operations
pub struct OfficeApi;

impl OfficeApi {
    /// Initialize Office.js and verify we're running in Excel
    pub async fn initialize() -> bool {
        return initializeOffice();
    }

    /// Read the value from a specific cell (e.g., "A1", "B2")
    pub async fn get_cell_value(address: &str) -> Result<String, String> {
        let promise = getCellValue(address);
        let future = JsFuture::from(promise);

        match future.await {
            Ok(value) => {
                let result = value.as_string().unwrap_or_default();
                console::log_1(&format!("Read cell {}: {}", address, result).into());
                Ok(result)
            }
            Err(err) => {
                let error_msg = format!("Failed to read cell {}: {:?}", address, err);
                console::log_1(&error_msg.clone().into());
                Err(error_msg)
            }
        }
    }

    /// Write a value to a specific cell
    pub async fn set_cell_value(address: &str, value: &str) -> Result<(), String> {
        let promise = setCellValue(address, value);
        let future = JsFuture::from(promise);

        match future.await {
            Ok(_) => {
                console::log_1(&format!("Set cell {} to: {}", address, value).into());
                Ok(())
            }
            Err(err) => {
                let error_msg = format!("Failed to write to cell {}: {:?}", address, err);
                console::log_1(&error_msg.clone().into());
                Err(error_msg)
            }
        }
    }

    /// Get the name of the active worksheet
    pub async fn get_active_sheet_name() -> Result<String, String> {
        let promise = getActiveSheetName();
        let future = JsFuture::from(promise);

        match future.await {
            Ok(value) => {
                let result = value.as_string().unwrap_or("Unknown".to_string());
                console::log_1(&format!("Active sheet: {}", result).into());
                Ok(result)
            }
            Err(err) => {
                let error_msg = format!("Failed to get sheet name: {:?}", err);
                console::log_1(&error_msg.clone().into());
                Err(error_msg)
            }
        }
    }
}
