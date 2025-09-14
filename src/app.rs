use crate::office_bindings::OfficeApi;
use futures::StreamExt;
use gloo_timers::future::IntervalStream;
use std::sync::{Arc, Mutex};
use wasm_bindgen_futures::spawn_local;
use web_sys::console;

pub struct OfficeState {
    office_status: bool,
    value: String,
    address: String,
    sheet: String,

    new_value: String,
}

/// We derive Deserialize/Serialize so we can persist app state on shutdown.
#[derive(serde::Deserialize, serde::Serialize)]
#[serde(default)] // if we add new fields, give them default values when deserializing old state
pub struct EGUIOfficeJsExcel {
    #[serde(skip)]
    office_state: Arc<Mutex<OfficeState>>,
}

impl Default for EGUIOfficeJsExcel {
    fn default() -> Self {
        Self {
            office_state: Arc::new(Mutex::new(OfficeState {
                office_status: false,
                value: "".to_owned(),
                address: "A1".to_owned(),
                sheet: "".to_owned(),
                new_value: "".to_owned(),
            })),
        }
    }
}

impl EGUIOfficeJsExcel {
    /// Called once before the first frame.
    pub fn new(cc: &eframe::CreationContext<'_>) -> Self {
        // This is also where you can customize the look and feel of egui using
        // `cc.egui_ctx.set_visuals` and `cc.egui_ctx.set_fonts`.

        // Load previous app state (if any).
        // Note that you must enable the `persistence` feature for this to work.
        let mut app: Self = if let Some(storage) = cc.storage {
            eframe::get_value(storage, eframe::APP_KEY).unwrap_or_default()
        } else {
            Default::default()
        };

        // Update office availability on startup and check readiness
        app.wait_for_office(&cc.egui_ctx);
        app
    }

    fn wait_for_office(&mut self, ctx: &egui::Context) {
        let ctx = ctx.clone();
        let office_state = self.office_state.clone();

        spawn_local(async move {
            let mut interval = IntervalStream::new(1000);
            let mut office_ready = OfficeApi::initialize().await;

            while !office_ready {
                interval.next().await;
                office_ready = OfficeApi::initialize().await;
            }
            (*office_state.lock().unwrap()).office_status = office_ready;
            console::log_1(&format!("office initialized.").into());
            ctx.request_repaint();
        });
    }

    fn read_cell(&mut self, ctx: &egui::Context) {
        let ctx = ctx.clone();
        let office_state = self.office_state.clone();

        if !(*office_state.lock().unwrap()).office_status {
            return;
        }

        let address = (*office_state.lock().unwrap()).address.clone();

        spawn_local(async move {
            match OfficeApi::get_cell_value(&address).await {
                Ok(value) => {
                    console::log_1(
                        &format!("Successfully read cell {}: {}", address, value).into(),
                    );
                    (*office_state.lock().unwrap()).value = value;
                }
                Err(e) => {
                    console::log_1(&format!("Failed to read cell {}: {}", address, e).into());
                }
            }
            ctx.request_repaint();
        });
    }

    fn write_cell(&mut self, ctx: &egui::Context) {
        let ctx = ctx.clone();
        let office_state = self.office_state.clone();

        if !(*office_state.lock().unwrap()).office_status {
            return;
        }

        let address = (*office_state.lock().unwrap()).address.clone();
        let value = (*office_state.lock().unwrap()).new_value.clone();

        spawn_local(async move {
            match OfficeApi::set_cell_value(&address, &value).await {
                Ok(_) => {
                    console::log_1(
                        &format!("Successfully wrote to cell {}: {}", address, value).into(),
                    );
                }
                Err(e) => {
                    console::log_1(
                        &format!("Failed to write to cell {}: {}\n{}", address, value, e).into(),
                    );
                }
            }
            ctx.request_repaint();
        });
    }

    fn get_sheet_name(&mut self, ctx: &egui::Context) {
        let ctx = ctx.clone();
        let office_state = self.office_state.clone();

        if !(*office_state.lock().unwrap()).office_status {
            return;
        }

        spawn_local(async move {
            match OfficeApi::get_active_sheet_name().await {
                Ok(name) => {
                    console::log_1(&format!("Successfully read sheet name: {}", name).into());
                    (*office_state.lock().unwrap()).sheet = name;
                }
                Err(e) => {
                    console::log_1(&format!("Failed to get sheet name: {}", e).into());
                }
            }
            ctx.request_repaint();
        });
    }
}

impl eframe::App for EGUIOfficeJsExcel {
    /// Called by the framework to save state before shutdown.
    fn save(&mut self, storage: &mut dyn eframe::Storage) {
        eframe::set_value(storage, eframe::APP_KEY, self);
    }

    /// Called each time the UI needs repainting, which may be many times per second.
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        let office_state = self.office_state.clone();
        // Office status is managed through async callbacks

        egui::TopBottomPanel::top("top_panel").show(ctx, |ui| {
            egui::MenuBar::new().ui(ui, |ui| {
                let is_web = cfg!(target_arch = "wasm32");
                if !is_web {
                    ui.menu_button("File", |ui| {
                        if ui.button("Quit").clicked() {
                            ctx.send_viewport_cmd(egui::ViewportCommand::Close);
                        }
                    });
                    ui.add_space(16.0);
                }

                egui::widgets::global_theme_preference_buttons(ui);
            });
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("egui + Office.js Excel Integration");

            ui.separator();

            // Office.js Status Section
            ui.heading("Office.js Status");

            ui.horizontal(|ui| {
                if (*office_state.lock().unwrap()).office_status {
                    ui.colored_label(egui::Color32::GREEN, "✓ Loaded");
                } else {
                    ui.colored_label(egui::Color32::RED, "✗ Not Loaded");
                }
            });

            ui.separator();

            // Cell Operations Section
            ui.heading("Cell Operations");

            ui.horizontal(|ui| {
                ui.label("Cell Address:");
                ui.text_edit_singleline(&mut office_state.lock().unwrap().address);
                ui.label("(e.g., A1, B2, C10)");
            });

            ui.horizontal(|ui| {
                ui.label("Current Value:");
                ui.text_edit_singleline(&mut office_state.lock().unwrap().value);
                if ui.button("Read Cell").clicked() {
                    self.read_cell(ctx);
                }
            });

            ui.horizontal(|ui| {
                ui.label("New Value:");
                ui.text_edit_singleline(&mut office_state.lock().unwrap().new_value);
                if ui.button("Write Cell").clicked() {
                    self.write_cell(ctx);
                }
            });

            ui.separator();

            // Sheet Information
            ui.heading("Sheet Information");

            ui.horizontal(|ui| {
                ui.label("Active Sheet:");
                ui.label(&office_state.lock().unwrap().sheet);
                if ui.button("Get Sheet Name").clicked() {
                    self.get_sheet_name(ctx);
                }
            });

            ui.separator();

            ui.with_layout(egui::Layout::bottom_up(egui::Align::LEFT), |ui| {
                ui.horizontal(|ui| {
                    ui.spacing_mut().item_spacing.x = 0.0;
                    ui.label("Powered by ");
                    ui.hyperlink_to("egui", "https://github.com/emilk/egui");
                    ui.label(", ");
                    ui.hyperlink_to(
                        "eframe",
                        "https://github.com/emilk/egui/tree/master/crates/eframe",
                    );
                    ui.label(", and ");
                    ui.hyperlink_to(
                        "Office.js",
                        "https://docs.microsoft.com/en-us/office/dev/add-ins/",
                    );
                    ui.label(".");
                });
                egui::warn_if_debug_build(ui);
            });
        });
    }
}
