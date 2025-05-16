📦 Tracciato Converter Web App

Una Web App sviluppata in Python con Streamlit per convertire file .txt o .zip contenenti record a lunghezza fissa in un file Excel strutturato e analizzabile. L'applicazione è pensata per gestire processi di tracciamento e controllo, con funzionalità di parsing, validazione e reportistica avanzata.
🚀 Funzionalità principali

    ✅ Supporto per file .txt e .zip

    🔍 Parsing dei record a lunghezza fissa

    📤 Esportazione in Excel con i seguenti fogli:

        Dati Fat: dati validi, con origine file e numero riga

        Scarti Fat: record scartati con motivazione

        Riepilogo Fat: riepilogo per file caricato

    📊 Grafici dinamici:

        Istogramma settimanale per Data Calc

        Top 10 destinatari per Peso e Importo

        Trend temporale

📁 Struttura del progetto

📦 app_marcopj/
├── app_marcopj.py            # Script principale Streamlit
├── requirements.txt          # Librerie richieste
└── README.md                 # (Questo file)

🧰 Requisiti

    Python ≥ 3.9

    Librerie elencate in requirements.txt

⚙️ Installazione (Windows/macOS/Linux)
1. Clona il repository

git clone https://github.com/tuo-utente/app_marcopj.git
cd app_marcopj

2. Installa le dipendenze

pip install -r requirements.txt

    Se usi un Mac con Miniconda:

conda create -n marcopj python=3.11
conda activate marcopj
pip install -r requirements.txt

🖥️ Avvio dell'applicazione

streamlit run app_marcopj.py

📌 Note aggiuntive

    L'app mostra anteprime dei dati validi e scartati direttamente nella UI.

    I file .zip possono contenere più file .txt; verranno processati tutti.

    Tutti i file Excel generati contengono tracking dettagliato: nome file, riga di origine, motivazioni di scarto.

📝 Licenza

Distribuito con licenza MIT. Vedi il file LICENSE per maggiori dettagli.
