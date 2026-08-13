import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import re
import io
from pathlib import Path

# Selenium is imported only when needed so the dashboard can still open
# if Selenium/Chrome is not installed.
try:
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    SELENIUM_AVAILABLE = True
except ImportError:
    SELENIUM_AVAILABLE = False

try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# ============================================================
# CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Exam Data Collection",
    page_icon="📊",
    layout="wide"
)

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "exam_data.db"

BOOKS_FILE = BASE_DIR / "books-toscrape.xlsx"
GAARAAS_FILE = BASE_DIR / "gaaraas-com.xlsx"

KOBO_URL = "https://ee.kobotoolbox.org/i/nhtuQWNB"
GOOGLE_FORM_URL = (
    "https://docs.google.com/forms/d/e/"
    "1FAIpQLScIQcbr1OyyZfgDLm4N1C-zCwo-Ck-ch6OIP_AVHC4m7APn1w/"
    "viewform?usp=publish-editor"
)

BOOKS_START_URL = "https://books.toscrape.com/catalogue/page-1.html"
GAARAAS_START_URL = "https://www.gaaraas.com/fr/"

# ============================================================
# GENERAL FUNCTIONS
# ============================================================

def clean_column_names(df):
    """Normalize column names."""
    df = df.copy()
    df.columns = [
        re.sub(r"[^a-zA-Z0-9_]+", "_", str(c).strip().lower()).strip("_")
        for c in df.columns
    ]
    return df


def clean_numeric(value):
    """Convert a value such as '£51.77', '5 500 000 F CFA', etc. to float."""
    if pd.isna(value):
        return np.nan
    text = str(value).replace(",", ".")
    text = re.sub(r"[^\d.\-]", "", text)
    if text in ("", ".", "-", "-."):
        return np.nan
    try:
        return float(text)
    except ValueError:
        return np.nan


def clean_dataframe(df, source):
    """Basic cleaning used before the dashboard and SQL storage."""
    df = clean_column_names(df)
    df = df.drop_duplicates().reset_index(drop=True)

    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace(
                {"nan": np.nan, "None": np.nan, "": np.nan}
            )

    # Convert obvious numeric columns.
    numeric_keywords = [
        "price", "prix", "rating", "star", "review", "reviews",
        "tax", "taxe", "number", "nombre", "mileage", "kilometrage",
        "kilométrage", "km", "year", "annee", "année"
    ]

    for col in df.columns:
        if any(key in col for key in numeric_keywords):
            converted = df[col].apply(clean_numeric)
            # Only replace when conversion produced useful numeric values.
            if converted.notna().sum() >= max(1, int(len(df) * 0.30)):
                df[col] = converted

    df["source"] = source
    return df


# ============================================================
# SQLITE
# ============================================================

def get_connection():
    return sqlite3.connect(DB_PATH)


def save_to_sql(df, table_name):
    """Store a dataframe in SQLite."""
    if df is None or df.empty:
        return 0

    data = df.copy()
    # SQLite cannot directly store complex objects.
    for col in data.columns:
        data[col] = data[col].apply(
            lambda x: x if pd.isna(x) or isinstance(x, (str, int, float, bool))
            else str(x)
        )

    with get_connection() as conn:
        data.to_sql(table_name, conn, if_exists="replace", index=False)

    return len(data)


def read_sql_table(table_name):
    with get_connection() as conn:
        tables = pd.read_sql_query(
            "SELECT name FROM sqlite_master WHERE type='table'",
            conn
        )["name"].tolist()

        if table_name not in tables:
            return pd.DataFrame()

        return pd.read_sql_query(f'SELECT * FROM "{table_name}"', conn)


def initialize_database():
    """Create the database and two source tables if they do not exist."""
    with get_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS books_toscrape (
                id INTEGER PRIMARY KEY AUTOINCREMENT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS gaaraas (
                id INTEGER PRIMARY KEY AUTOINCREMENT
            )
        """)
        conn.commit()


# ============================================================
# EXCEL DOWNLOADS
# ============================================================

def dataframe_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return output.getvalue()


def load_excel_file(path):
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path)
    except Exception as exc:
        st.error(f"Impossible de lire {path.name}: {exc}")
        return pd.DataFrame()


# ============================================================
# SELENIUM
# ============================================================

def create_driver():
    if not SELENIUM_AVAILABLE:
        raise RuntimeError(
            "Selenium n'est pas installé. Lancez : pip install selenium"
        )

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 Chrome/131 Safari/537.36"
    )

    # Selenium Manager normally detects/downloads a compatible driver.
    return webdriver.Chrome(options=options)


def scrape_books_toscrape(nb_pages=5):
    """
    Scrape Books to Scrape over several pages with Selenium.

    This follows the usual Books to Scrape catalogue structure:
    article.product_pod, h3 a, .price_color, .instock.availability,
    .star-rating and the product detail URL.
    """
    driver = create_driver()
    rows = []

    try:
        for page in range(1, nb_pages + 1):
            url = f"https://books.toscrape.com/catalogue/page-{page}.html"
            driver.get(url)

            WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, "article.product_pod")
                )
            )

            cards = driver.find_elements(By.CSS_SELECTOR, "article.product_pod")

            for card in cards:
                try:
                    title_el = card.find_element(By.CSS_SELECTOR, "h3 a")
                    title = title_el.get_attribute("title") or title_el.text.strip()
                except Exception:
                    title = None

                try:
                    price = card.find_element(
                        By.CSS_SELECTOR, ".price_color"
                    ).text.strip()
                except Exception:
                    price = None

                try:
                    availability = card.find_element(
                        By.CSS_SELECTOR, ".availability"
                    ).text.strip()
                except Exception:
                    availability = None

                try:
                    rating_class = card.find_element(
                        By.CSS_SELECTOR, "p.star-rating"
                    ).get_attribute("class")
                    rating = rating_class.replace("star-rating", "").strip()
                except Exception:
                    rating = None

                try:
                    relative_url = title_el.get_attribute("href")
                    product_url = relative_url
                except Exception:
                    product_url = None

                rows.append({
                    "title": title,
                    "price": price,
                    "availability": availability,
                    "star_rating": rating,
                    "book_url": product_url,
                    "page": page
                })
    finally:
        driver.quit()

    return pd.DataFrame(rows)


def first_text(element, selectors):
    """Return the first non-empty text found with one of the selectors."""
    for selector in selectors:
        try:
            value = element.find_element(By.CSS_SELECTOR, selector).text.strip()
            if value:
                return value
        except Exception:
            pass
    return None


def scrape_gaaraas(nb_pages=3):
    """
    Selenium scraper for Gaaraas.

    Gaaraas has changed its front-end structure over time. Therefore,
    several CSS selectors are tried for each field. If the live HTML
    changes, update the selector lists below to match the current
    ExamDataCollection.ipynb selectors.
    """
    driver = create_driver()
    rows = []

    try:
        for page in range(1, nb_pages + 1):
            # Common pagination patterns. The first page uses the root URL.
            if page == 1:
                url = GAARAAS_START_URL
            else:
                url = f"{GAARAAS_START_URL}?page={page}"

            driver.get(url)

            # Wait for the page to render.
            WebDriverWait(driver, 15).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )

            # Try several possible listing-card selectors.
            selectors = [
                "article",
                ".vehicle-card",
                ".car-card",
                ".listing-card",
                ".item",
                "[class*='vehicle']",
                "[class*='listing']",
                "[class*='card']"
            ]

            cards = []
            for selector in selectors:
                try:
                    found = driver.find_elements(By.CSS_SELECTOR, selector)
                    if len(found) >= 3:
                        cards = found
                        break
                except Exception:
                    pass

            if not cards:
                # Fallback: collect links that appear to be listings.
                links = driver.find_elements(By.CSS_SELECTOR, "a[href]")
                for link in links:
                    href = link.get_attribute("href") or ""
                    text = link.text.strip()
                    if text and (
                        "gaaraas.com" in href
                        and ("/fr/" in href or "/en/" in href)
                    ):
                        rows.append({
                            "title": text,
                            "price": None,
                            "location": None,
                            "details": None,
                            "listing_url": href,
                            "page": page
                        })
                continue

            for card in cards:
                title = first_text(
                    card,
                    [
                        "h2", "h3", "h4",
                        ".title", ".name",
                        "[class*='title']", "[class*='name']"
                    ]
                )

                price = first_text(
                    card,
                    [
                        ".price",
                        "[class*='price']",
                        "[class*='prix']"
                    ]
                )

                location = first_text(
                    card,
                    [
                        ".location",
                        "[class*='location']",
                        "[class*='local']"
                    ]
                )

                details = first_text(
                    card,
                    [
                        ".details",
                        ".description",
                        "[class*='detail']",
                        "[class*='description']"
                    ]
                )

                try:
                    link = card.find_element(By.CSS_SELECTOR, "a[href]")
                    listing_url = link.get_attribute("href")
                except Exception:
                    listing_url = None

                if any([title, price, location, details, listing_url]):
                    rows.append({
                        "title": title,
                        "price": price,
                        "location": location,
                        "details": details,
                        "listing_url": listing_url,
                        "page": page
                    })
    finally:
        driver.quit()

    df = pd.DataFrame(rows)

    # Remove repeated cards generated by multiple selector fallbacks.
    if not df.empty:
        subset = [c for c in ["title", "listing_url", "price"] if c in df.columns]
        if subset:
            df = df.drop_duplicates(subset=subset)

    return df


# ============================================================
# DASHBOARD
# ============================================================

def show_dashboard(df, title):
    st.subheader(title)

    if df.empty:
        st.warning("Aucune donnée disponible.")
        return

    df = clean_column_names(df)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Lignes", len(df))
    c2.metric("Colonnes", len(df.columns))
    c3.metric("Doublons", int(df.duplicated().sum()))
    c4.metric("Valeurs manquantes", int(df.isna().sum().sum()))

    st.dataframe(df, use_container_width=True)

    if not PLOTLY_AVAILABLE:
        st.info("Installez plotly pour afficher les graphiques.")
        return

    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    text_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()

    if numeric_cols:
        st.markdown("### Distribution des variables numériques")
        selected_num = st.selectbox(
            "Choisir une variable",
            numeric_cols,
            key=f"num_{title}"
        )
        fig = px.histogram(
            df,
            x=selected_num,
            title=f"Distribution de {selected_num}"
        )
        st.plotly_chart(fig, use_container_width=True)

    if text_cols:
        st.markdown("### Analyse catégorielle")
        selected_text = st.selectbox(
            "Choisir une variable catégorielle",
            text_cols,
            key=f"text_{title}"
        )

        counts = (
            df[selected_text]
            .value_counts(dropna=False)
            .head(15)
            .reset_index()
        )
        counts.columns = [selected_text, "count"]

        fig = px.bar(
            counts,
            x=selected_text,
            y="count",
            title=f"Top valeurs - {selected_text}"
        )
        st.plotly_chart(fig, use_container_width=True)


# ============================================================
# APP
# ============================================================

initialize_database()

st.title("📊 Exam Data Collection")
st.caption(
    "Application Streamlit — Selenium, Web Scraper, Dashboard, "
    "formulaires d'évaluation et stockage SQL"
)

with st.sidebar:
    st.header("Navigation")
    page = st.radio(
        "Choisir une rubrique",
        [
            "🏠 Accueil",
            "🕷️ Scraping Selenium",
            "📥 Données brutes Web Scraper",
            "📊 Dashboard",
            "📝 Évaluation",
            "🗄️ Base SQL"
        ]
    )

# ------------------------------------------------------------
# HOME
# ------------------------------------------------------------
if page == "🏠 Accueil":
    st.markdown("## Projet ExamDataCollection")

    st.info(
        "Repository GitHub : "
        "https://github.com/AlyCoto/ExamDataCollection"
    )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### Sources Selenium")
        st.markdown("- Books to Scrape")
        st.markdown("- Gaaraas")

    with col2:
        st.markdown("### Fonctions de l'application")
        st.markdown("- Scraping multi-pages")
        st.markdown("- Téléchargement des XLSX Web Scraper")
        st.markdown("- Dashboard des données nettoyées")
        st.markdown("- Formulaires Kobo + Google Forms")
        st.markdown("- Base SQL SQLite")

# ------------------------------------------------------------
# SELENIUM
# ------------------------------------------------------------
elif page == "🕷️ Scraping Selenium":
    st.header("🕷️ Scraping avec Selenium")

    st.write(
        "Cette partie reprend le principe du notebook "
        "`ExamDataCollection.ipynb` et permet de lancer le scraping "
        "sur plusieurs pages directement depuis Streamlit."
    )

    if not SELENIUM_AVAILABLE:
        st.error(
            "Selenium n'est pas disponible dans cet environnement. "
            "Installez les dépendances avec le fichier requirements.txt."
        )
    else:
        source = st.selectbox(
            "Source à scraper",
            ["Books to Scrape", "Gaaraas"]
        )

        default_pages = 5 if source == "Books to Scrape" else 3
        nb_pages = st.number_input(
            "Nombre de pages",
            min_value=1,
            max_value=100,
            value=default_pages,
            step=1
        )

        if source == "Books to Scrape":
            st.code(BOOKS_START_URL)
        else:
            st.code(GAARAAS_START_URL)

        if st.button("🚀 Lancer le scraping", type="primary"):
            try:
                with st.spinner("Scraping en cours..."):
                    if source == "Books to Scrape":
                        raw_df = scrape_books_toscrape(int(nb_pages))
                        cleaned_df = clean_dataframe(raw_df, "Books to Scrape")
                        table = "books_toscrape"
                    else:
                        raw_df = scrape_gaaraas(int(nb_pages))
                        cleaned_df = clean_dataframe(raw_df, "Gaaraas")
                        table = "gaaraas"

                st.session_state[f"raw_{source}"] = raw_df
                st.session_state[f"clean_{source}"] = cleaned_df

                if cleaned_df.empty:
                    st.warning(
                        "Le scraping n'a retourné aucune ligne. "
                        "Pour Gaaraas, vérifiez les CSS selectors du site "
                        "et ceux utilisés dans votre notebook."
                    )
                else:
                    save_to_sql(cleaned_df, table)
                    st.success(
                        f"{len(cleaned_df)} lignes récupérées et enregistrées "
                        f"dans la table SQL `{table}`."
                    )

                    st.dataframe(
                        cleaned_df,
                        use_container_width=True
                    )

                    st.download_button(
                        "⬇️ Télécharger les données Selenium",
                        data=dataframe_to_excel(cleaned_df),
                        file_name=(
                            "books_toscrape_selenium.xlsx"
                            if source == "Books to Scrape"
                            else "gaaraas_selenium.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        )
                    )

# ------------------------------------------------------------
# RAW WEB SCRAPER DATA
# ------------------------------------------------------------
elif page == "📥 Données brutes Web Scraper":
    st.header("📥 Données brutes issues du Web Scraper")

    st.write(
        "Les deux fichiers XLSX présents dans le repository sont proposés "
        "au téléchargement."
    )

    files = [
        ("Books to Scrape", BOOKS_FILE),
        ("Gaaraas", GAARAAS_FILE)
    ]

    for label, path in files:
        st.markdown(f"### {label}")

        df = load_excel_file(path)

        if df.empty:
            st.warning(
                f"Le fichier `{path.name}` est introuvable ou vide. "
                "Placez-le dans le même dossier que myapp.py."
            )
        else:
            st.write(f"{len(df)} lignes — {len(df.columns)} colonnes")
            st.dataframe(df.head(10), use_container_width=True)

            st.download_button(
                f"⬇️ Télécharger {path.name}",
                data=path.read_bytes(),
                file_name=path.name,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key=f"download_{path.name}"
            )

# ------------------------------------------------------------
# DASHBOARD
# ------------------------------------------------------------
elif page == "📊 Dashboard":
    st.header("📊 Dashboard des données nettoyées")

    source = st.selectbox(
        "Choisir la source",
        ["Books to Scrape", "Gaaraas"]
    )

    key = f"clean_{source}"

    if key in st.session_state and not st.session_state[key].empty:
        df = st.session_state[key]
    else:
        # Fall back to the corresponding Web Scraper XLSX.
        if source == "Books to Scrape":
            df = load_excel_file(BOOKS_FILE)
        else:
            df = load_excel_file(GAARAAS_FILE)

        if not df.empty:
            df = clean_dataframe(df, source)

    show_dashboard(
        df,
        f"Dashboard — {source}"
    )

# ------------------------------------------------------------
# FORMS
# ------------------------------------------------------------
elif page == "📝 Évaluation":
    st.header("📝 Évaluation de l'application")

    st.markdown(
        "Votre retour nous permet d'évaluer l'application et "
        "d'identifier les points à améliorer."
    )

    tab1, tab2 = st.tabs(["KoboToolbox", "Google Forms"])

    with tab1:
        st.subheader("Formulaire KoboToolbox")
        st.components.v1.html(
            f"""
            <iframe
                src="{KOBO_URL}"
                width="100%"
                height="700"
                frameborder="0">
            </iframe>
            """,
            height=720,
            scrolling=True
        )
        st.link_button(
            "Ouvrir KoboToolbox dans un nouvel onglet",
            KOBO_URL
        )

    with tab2:
        st.subheader("Formulaire Google Forms")
        st.components.v1.html(
            f"""
            <iframe
                src="{GOOGLE_FORM_URL}"
                width="100%"
                height="700"
                frameborder="0">
            </iframe>
            """,
            height=720,
            scrolling=True
        )
        st.link_button(
            "Ouvrir Google Forms dans un nouvel onglet",
            GOOGLE_FORM_URL
        )

# ------------------------------------------------------------
# SQL
# ------------------------------------------------------------
elif page == "🗄️ Base SQL":
    st.header("🗄️ Base de données SQL")

    st.write(
        "L'application utilise SQLite. Deux tables sont prévues, une "
        "pour chaque source de données : `books_toscrape` et `gaaraas`."
    )

    st.code(
        """
CREATE TABLE books_toscrape (...);
CREATE TABLE gaaraas (...);
        """,
        language="sql"
    )

    tables = ["books_toscrape", "gaaraas"]

    for table in tables:
        df_sql = read_sql_table(table)

        if not df_sql.empty:
            st.subheader(f"Table : {table}")
            st.write(f"{len(df_sql)} lignes")
            st.dataframe(df_sql, use_container_width=True)
        else:
            st.info(
                f"La table `{table}` est vide. Lancez le scraping Selenium "
                "pour l'alimenter."
            )

    st.markdown("---")
    st.write(f"Fichier SQLite : `{DB_PATH.name}`")

    if DB_PATH.exists():
        st.download_button(
            "⬇️ Télécharger la base SQLite",
            data=DB_PATH.read_bytes(),
            file_name="exam_data.db",
            mime="application/octet-stream"
        )

st.markdown("---")
st.caption("ExamDataCollection — Streamlit / Selenium / SQLite")
