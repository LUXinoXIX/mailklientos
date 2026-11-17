// Prototyp MailKlienta na Tauri 2: rovnaké rozhranie (React) v okne WebView2 (jadro Microsoft Edge).
// Zatiaľ len ukážkové dáta – slúži na overenie, či na danom počítači zmizne blikanie.
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
    tauri::Builder::default()
        .run(tauri::generate_context!())
        .expect("aplikáciu sa nepodarilo spustiť");
}
