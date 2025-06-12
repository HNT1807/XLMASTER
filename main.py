import streamlit as st
import pandas as pd
import os
import io
import zipfile
import base64
import time
import uuid
import re

# from openpyxl.utils import get_column_letter # Not strictly needed if using cell.column_letter

# --- Configuration (Using Excel Column Letters) ---
FILENAME_COLUMN_LETTER = 'B'
TRACK_TITLE_COLUMN_LETTER = 'R'

EXCLUDED_COLUMNS_LETTERS = [
    'A', 'B', 'C', 'D', 'E', 'K', 'S', 'T', 'U', 'V', 'X', 'Y',
    'AE', 'AI', 'AP', 'BC', 'BD'  # AI is rule-defined but based on source AI
]


# --- End of Configuration ---

# --- Helper Functions ---
def excel_col_to_index(col_str):
    if not isinstance(col_str, str) or not col_str.isalpha():
        raise ValueError(f"Invalid Excel column letter: {col_str}")
    index = 0;
    col_str = col_str.upper()
    for char in col_str: index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def index_to_excel_col(n):
    n_orig = n;
    n += 1;
    string = ""
    if n <= 0: return f"InvalidIndex({n_orig})"
    while n > 0: n, rem = divmod(n - 1, 26); string = chr(65 + rem) + string
    return string


try:
    FILENAME_COL_IDX = excel_col_to_index(FILENAME_COLUMN_LETTER)
    TRACK_TITLE_COL_IDX = excel_col_to_index(TRACK_TITLE_COLUMN_LETTER)
    EXCLUDED_COL_INDICES = sorted(list(set(excel_col_to_index(c) for c in EXCLUDED_COLUMNS_LETTERS)))
    A_IDX = excel_col_to_index('A');
    C_IDX = excel_col_to_index('C');
    E_IDX = excel_col_to_index('E')
    K_IDX = excel_col_to_index('K');
    P_IDX = excel_col_to_index('P');
    S_IDX = excel_col_to_index('S')
    T_IDX = excel_col_to_index('T');
    U_IDX = excel_col_to_index('U');
    AE_IDX = excel_col_to_index('AE')
    AI_IDX = excel_col_to_index('AI');
    BC_IDX = excel_col_to_index('BC');
    BD_IDX = excel_col_to_index('BD')
    V_IDX = excel_col_to_index('V');
    Y_IDX = excel_col_to_index('Y')
except ValueError as e:
    st.error(f"Configuration Error in column letters: {e}"); st.stop()


def get_raw_stem_part_from_filename(fn_str):
    if not fn_str or not isinstance(fn_str, str): return None
    name, _ = os.path.splitext(fn_str);
    m = re.search(r"_STEM(.*)", name)
    return m.group(1) if m and m.group(1) else None


def format_extracted_stem_part(stem_raw):
    if not stem_raw or not isinstance(stem_raw, str): return ""
    txt = str(stem_raw)
    txt = re.sub(r"([a-z\d])([A-Z])", r"\1 \2", txt)
    txt = re.sub(r"([A-Z])([A-Z][a-z])", r"\1 \2", txt)
    txt = re.sub(r"([A-Za-z])(\d)", r"\1 \2", txt)
    return re.sub(r'\s+', ' ', txt).strip()


def extract_main_title_from_filename_robust(fn_str):
    if not isinstance(fn_str, str) or not fn_str.strip(): return None
    try:
        name_without_ext, _ = os.path.splitext(fn_str)
        name_for_title_extraction = name_without_ext
        suffixes_to_remove = ["_STEM", "_Full"]
        for suffix_base in suffixes_to_remove:
            if suffix_base in name_for_title_extraction:
                name_for_title_extraction = name_for_title_extraction.split(suffix_base)[0]
                break
        parts = name_for_title_extraction.split('_')
        if len(parts) >= 3:
            title = "_".join(parts[2:]); return title.strip() if title else None
        elif len(parts) == 2:
            return parts[1].strip() if parts[1] else None
        elif len(parts) == 1 and name_for_title_extraction.strip():
            return name_for_title_extraction.strip()
    except Exception:
        return None
    return None


def get_track_number_from_filename(fn_str):
    if not isinstance(fn_str, str) or not fn_str.strip(): return ""
    name_without_ext, _ = os.path.splitext(str(fn_str));
    parts = name_without_ext.split('_')
    if len(parts) >= 2: return parts[1]
    return ""


def get_col_E_value_from_filename(fn_str):
    if not isinstance(fn_str, str) or not fn_str.strip(): return ""
    name_without_ext, _ = os.path.splitext(str(fn_str));
    parts = name_without_ext.split('_')
    if len(parts) >= 2:
        return "_".join(parts[:2])
    elif len(parts) == 1:
        return parts[0]
    return ""


def auto_adjust_column_width(worksheet):
    for col in worksheet.columns:
        max_length = 0;
        column_letter = col[0].column_letter
        try:  # Consider header length
            header_val = worksheet[f"{column_letter}1"].value
            if header_val: max_length = len(str(header_val))
        except:
            pass

        for cell in col:
            try:
                if cell.value:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length: max_length = cell_length
            except:
                pass
        adjusted_width = (max_length + 2)
        if adjusted_width > 70: adjusted_width = 70  # Cap width
        worksheet.column_dimensions[column_letter].width = adjusted_width


def trigger_download_component(bin_file, download_filename, mime_type):
    b64 = base64.b64encode(bin_file).decode();
    script_id = f"dl_script_{uuid.uuid4().hex}"
    return f"""<html><head><meta charset="UTF-8"></head><body><script id="{script_id}">
        (function(){{var e=document.createElement('a');e.setAttribute('href','data:{mime_type};base64,{b64}');
        e.setAttribute('download','{download_filename}');e.style.display='none';document.body.appendChild(e);
        try{{e.click();}}catch(err){{console.error('Download click error for {download_filename}:',err);}}
        finally{{document.body.removeChild(e);}}}})();</script></body></html>"""


st.set_page_config(layout="wide");
st.title("XL MASTER")
st.markdown(f"Upload Excel files to batch process them. Downloads will start automatically.")
uploaded_files = st.file_uploader("Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True)
status_area = st.container();
download_trigger_area = st.container()

if uploaded_files:
    if st.button("🚀 Process Files and Auto-Download"):
        processed_data_outputs = [];
        skipped_files_count = 0;
        processed_files_count = 0
        with status_area:
            st.info("Processing files...");
            overall_progress_bar = st.progress(0);
            current_file_status = st.empty()
            for i, uploaded_file_obj in enumerate(uploaded_files):
                current_file_status.info(f"Processing: {uploaded_file_obj.name} ({i + 1}/{len(uploaded_files)})")
                current_df_shape = (0, 0)
                try:
                    df_original = pd.read_excel(uploaded_file_obj, engine='openpyxl' if uploaded_file_obj.name.endswith(
                        '.xlsx') else 'xlrd', header=0)
                    current_df_shape = df_original.shape

                    df_processed = df_original.copy();
                    file_was_modified = False

                    if A_IDX < current_df_shape[1]:
                        ca_val = 1
                        for r_i in range(df_processed.shape[0]):
                            if FILENAME_COL_IDX < current_df_shape[1] and pd.notna(
                                    df_processed.iloc[r_i, FILENAME_COL_IDX]) and str(
                                    df_processed.iloc[r_i, FILENAME_COL_IDX]).strip():
                                if not (pd.notna(df_processed.iloc[r_i, A_IDX]) and str(
                                        df_processed.iloc[r_i, A_IDX]) == str(ca_val)):
                                    df_processed.iloc[r_i, A_IDX] = ca_val;
                                    file_was_modified = True
                                ca_val += 1
                    if AE_IDX < current_df_shape[1]:
                        cae_val = 1
                        for r_i in range(df_processed.shape[0]):
                            if FILENAME_COL_IDX < current_df_shape[1] and pd.notna(
                                    df_processed.iloc[r_i, FILENAME_COL_IDX]) and str(
                                    df_processed.iloc[r_i, FILENAME_COL_IDX]).strip():
                                if not (pd.notna(df_processed.iloc[r_i, AE_IDX]) and str(
                                        df_processed.iloc[r_i, AE_IDX]) == str(cae_val)):
                                    df_processed.iloc[r_i, AE_IDX] = cae_val;
                                    file_was_modified = True
                                cae_val += 1

                    source_title_map_for_generic_copy = {}
                    main_title_to_first_original_track_no_map = {}

                    for _, row_map_data in df_original.iterrows():
                        if TRACK_TITLE_COL_IDX < current_df_shape[1] and pd.notna(
                                row_map_data.iloc[TRACK_TITLE_COL_IDX]):
                            title_in_R = str(row_map_data.iloc[TRACK_TITLE_COL_IDX]).strip()
                            if title_in_R and title_in_R not in source_title_map_for_generic_copy:
                                source_title_map_for_generic_copy[title_in_R] = row_map_data

                        if FILENAME_COL_IDX < current_df_shape[1] and pd.notna(row_map_data.iloc[FILENAME_COL_IDX]):
                            filename_for_map = str(row_map_data.iloc[FILENAME_COL_IDX])
                            main_title_for_map = extract_main_title_from_filename_robust(filename_for_map)
                            main_title_for_map = main_title_for_map.strip() if main_title_for_map else ""

                            if main_title_for_map and main_title_for_map not in main_title_to_first_original_track_no_map:
                                original_track_no = get_track_number_from_filename(filename_for_map)
                                if original_track_no:
                                    main_title_to_first_original_track_no_map[main_title_for_map] = original_track_no

                    cols_to_copy = [ci for ci in range(df_processed.shape[1]) if
                                    ci not in EXCLUDED_COL_INDICES and ci != TRACK_TITLE_COL_IDX]

                    for row_idx, row_s in df_processed.iterrows():
                        if not (FILENAME_COL_IDX < current_df_shape[1]): continue
                        if not (FILENAME_COL_IDX < len(row_s)): continue

                        fn_b = row_s.iloc[FILENAME_COL_IDX]
                        tt_r = row_s.iloc[TRACK_TITLE_COL_IDX] if TRACK_TITLE_COL_IDX < len(row_s) else ""

                        # --- Step 1: Always extract info from filename for every row ---
                        if not (pd.notna(fn_b) and str(fn_b).strip()):
                            continue # Skip rows with no filename

                        main_tt_current_row = extract_main_title_from_filename_robust(str(fn_b))
                        main_tt_current_row = main_tt_current_row.strip() if main_tt_current_row else ""
                        raw_stem = get_raw_stem_part_from_filename(str(fn_b))
                        fmt_stem = format_extracted_stem_part(raw_stem)
                        is_vocal = "vocal" in fmt_stem.lower()
                        match_src_row_for_generic_copy = None # Reset for each row

                        # --- Step 2: Conditionally fill in MISSING data in Column R and copy other data ---
                        if main_tt_current_row and (pd.isna(tt_r) or str(tt_r).strip() == ""):
                            file_was_modified = True
                            
                            # A) Fill in the missing track title in Column R
                            df_processed.iloc[row_idx, TRACK_TITLE_COL_IDX] = main_tt_current_row
                            
                            # B) Find a source row to copy data from
                            if main_tt_current_row in source_title_map_for_generic_copy:
                                match_src_row_for_generic_copy = source_title_map_for_generic_copy[main_tt_current_row]
                                for ctc_idx in cols_to_copy:
                                    if ctc_idx < current_df_shape[1] and ctc_idx < len(match_src_row_for_generic_copy):
                                        df_processed.iloc[row_idx, ctc_idx] = match_src_row_for_generic_copy.iloc[ctc_idx]

                            # C) Fill Column V
                            if V_IDX < current_df_shape[1]:
                                if main_tt_current_row in main_title_to_first_original_track_no_map:
                                    df_processed.iloc[row_idx, V_IDX] = main_title_to_first_original_track_no_map[main_tt_current_row]

                        # --- Step 3: Populate columns for ALL STEM rows, outside the conditional block ---
                        # This section now runs for every row that has a filename.

                        # Populate Column Y (Instrumentation)
                        if Y_IDX < current_df_shape[1]:
                            INSTRUMENT_KEYWORD_MAP = {
                                "accordion": "Accordion", "alpenhorn": "Alpenhorn/Alpine Horn",
                                "alpine horn": "Alpenhorn/Alpine Horn",
                                "autoharp": "Autoharp", "bagpipes": "Bagpipes", "bajo sexto": "Bajo Sexto",
                                "balafon": "Balafon",
                                "balalaika": "Balalaika", "bandoneon": "Bandoneon", "bandura": "Bandura",
                                "banjo": "Banjo",
                                "bansuri": "Bansuri/Baanhi/Baashi/Bansi/Basari",
                                "baanhi": "Bansuri/Baanhi/Baashi/Bansi/Basari",
                                "baashi": "Bansuri/Baanhi/Baashi/Bansi/Basari",
                                "bansi": "Bansuri/Baanhi/Baashi/Bansi/Basari",
                                "basari": "Bansuri/Baanhi/Baashi/Bansi/Basari", "bass": "Bass",
                                "bass drum": "Bass Drum",  # Order might matter if stem is "Bass Drum"
                                "bassoon": "Bassoon", "batacada": "Batacada", "bawu": "Bawu", "bell tree": "Bell Tree",
                                "bells": "Bells", "berimbau": "Berimbau", "big band": "Big Band",
                                "bladder pipe": "Bladder Pipe",
                                "bodhran": "Bodhran/Frame Drum", "frame drum": "Bodhran/Frame Drum",
                                "bombard": "Bombard",
                                "bombo": "Bombo", "bones": "Bones", "bongos": "Bongos", "bottle": "Bottle",
                                "bouzouki": "Bouzouki",
                                "bow": "Bow", "brass": "Brass", "bugle": "Bugle", "bullroarer": "Bullroarer/Rhombus",
                                "rhombus": "Bullroarer/Rhombus", "cabasa": "Cabasa", "calliope": "Calliope",
                                "carillon": "Carillon",
                                "castanets": "Castanets", "cavaquinho": "Cavaquinho", "celeste": "Celeste",
                                "cello": "Cello",
                                "chapman stick": "Chapman Stick", "charango": "Charango", "chekere": "Chekere/Djabara",
                                "djabara": "Chekere/Djabara", "chimes": "Chimes/Tubular Bells",
                                "tubular bells": "Chimes/Tubular Bells",
                                "cimbalom": "Cimbalom", "cittern": "Cittern", "clarinet": "Clarinet",
                                "clarsach": "Clarsach",
                                "claves": "Claves", "clavinet": "Clavinet", "coconuts": "Coconuts",
                                "comb and paper": "Comb And Paper",
                                "concertina": "Concertina", "conch shell": "Conch Shell", "congas": "Congas",
                                "cor anglais": "Cor Anglais/English Horn", "english horn": "Cor Anglais/English Horn",
                                "cornamuse": "Cornamuse", "cornet": "Cornet", "cornett": "Cornett",
                                "cowbell": "Cowbell",
                                "crotales": "Crotales", "crowth": "Crowth", "crumhorn": "Crumhorn", "cuatro": "Cuatro",
                                "cuica": "Cuica", "cymbals": "Cymbals", "da suo": "Da Suo", "daf": "Daf/Dayereh",
                                "dayereh": "Daf/Dayereh", "dan bau": "Dan Bau", "darbouka": "Darbouka", "def": "Def",
                                "descant fiddle": "Descant Fiddle", "dhol": "Dhol", "dholak": "Dholak",
                                "didgeridoo": "Didgeridoo",
                                "dilruba": "Dilruba", "dizi": "Dizi", "djembe": "Djembe", "dolceola": "Dolceola",
                                "double bass": "Double Bass", "doumbek/dumbek": "Doumbek/Dumbek",
                                "doumbek": "Doumbek/Dumbek", "dumbek": "Doumbek/Dumbek",
                                "drone": "Drone", "drum kit": "Drum Kit",
                                "drum machine": "Drum Machine/Electronic Drums",
                                "electronic drums": "Drum Machine/Electronic Drums", "drum set": "Drum Set",
                                "drums": "Drums",  # Generic 'drums'
                                "duck call": "Duck Call", "dudak": "Dudak", "dudu": "Dudu", "duduk": "Duduk",
                                "duff": "Duff",
                                "dulcimer": "Dulcimer", "dulcitone": "Dulcitone", "dunun": "Dunun",
                                "electronic instruments": "Electronic Instruments", "electronics": "Electronics",
                                "erhu": "Erhu",
                                "esraj": "Esraj", "ethnic plucked instruments": "Ethnic Plucked Instruments",
                                "ethnic string instruments": "Ethnic String Instruments",
                                "ethnic wind instruments": "Ethnic Wind Instruments",
                                "fiddle": "Fiddle", "fife": "Fife", "finger bells": "Finger Cymbals/Finger Bells",
                                "finger cymbals": "Finger Cymbals/Finger Bells", "finger snaps": "Finger Snaps",
                                "flapamba": "Flapamba",
                                "flexatone": "Flexatone", "flugelhorn": "Flugelhorn", "flute": "Flute", "fue": "Fue",
                                "gambang": "Gambang", "gamelan": "Gamelan", "gemshorn": "Gemshorn",
                                "ghaychak": "Ghaychak",
                                "ghurzen": "Ghurzen", "glockenspiel": "Glockenspiel",
                                "goblet drum": "Goblet Drum/Dumbec",
                                "dumbec": "Goblet Drum/Dumbec", "gong": "Gong", "gong - chinese": "Gong - Chinese/Chau",
                                "chau": "Gong - Chinese/Chau", "gran cassa": "Gran Cassa", "guiro": "Guiro",
                                "acoustic guitars": "Guitar - Acoustic/Steel String",
                                "acoustic guitar": "Guitar - Acoustic/Steel String",
                                "guitar acoustic": "Guitar - Acoustic/Steel String",
                                "guitars acoustic": "Guitar - Acoustic/Steel String",
                                "guitar - acoustic": "Guitar - Acoustic/Steel String",
                                "guitar - distorted electric": "Guitar - Distorted Electric",
                                "dobro": "Guitar - Dobro", "e-bow": "Guitar - E-Bow",
                                "electric guitars": "Guitar - Electric",
                                "electric guitar": "Guitar - Electric",
                                "guitars electric": "Guitar - Electric",
                                "guitar electric": "Guitar - Electric",  # Might conflict with "Guitar - Electric"
                                "guitar - electric": "Guitar - Electric",  # More specific
                                "pedal steel": "Guitar - Pedal Steel", "guitarron": "Guitarron", "guqin": "Guqin",
                                "guzheng": "Guzheng",
                                "hammered dulcimer": "Hammered Dulcimer", "hand claps": "Hand Claps",
                                "hang drum": "Hang Drum",
                                "harmonica": "Harmonica", "harmonium": "Harmonium", "harp": "Harp",
                                "harpsichord": "Harpsichord",
                                "hi-hat": "Hi-Hat", "hi hat": "Hi-Hat", "hihat": "Hi-Hat", "hichiriki": "Hichiriki",
                                "horn": "Horn", "french horn": "Horn - French", "horns": "Horns/Horn Section",
                                "hurdy gurdy": "Hurdy Gurdy",
                                "jazz trio": "Jazz Trio", "jug": "Jug", "kalimba/sanza": "Kalimba/Sanza",
                                "kalimba": "Kalimba/Sanza", "sanza": "Kalimba/Sanza",
                                "kamancheh": "Kamancheh/Kamanche/Kamancha", "kamanche": "Kamancheh/Kamanche/Kamancha",
                                "kamancha": "Kamancheh/Kamanche/Kamancha", "kanun": "Kanun", "kaval": "Kaval",
                                "kawala": "Kawala/Salamiya", "salamiya": "Kawala/Salamiya", "kazoo": "Kazoo",
                                "kecapi": "Kecapi",
                                "keyboard": "Keyboard", "keys": "Keyboard", "khene": "Khene", "khlui": "Khlui",
                                "koboz": "Koboz",
                                "kokyu": "Kokyu", "kora": "Kora", "kortholt": "Kortholt", "koto": "Koto",
                                "llamas hooves": "Llamas Hooves", "log drum": "Log Drum", "lute": "Lute",
                                "lyre": "Lyre",
                                "mallet": "Mallet", "mandira": "Mandira", "mandocello": "Mandocello",
                                "mandola": "Mandola",
                                "mandolin": "Mandolin", "maracas": "Maracas", "marimba": "Marimba",
                                "marimbula": "Marimbula",
                                "matou qin/morin khuur/horsehead fiddle": "MaTou Qin/Morin Khuur/Horsehead Fiddle",
                                # This is very long, ensure it's a likely stem name part
                                "mbira": "Mbira/African Thumb Piano",
                                "african thumb piano": "Mbira/African Thumb Piano",
                                "mellophone": "Mellophone", "mellotron": "Mellotron", "melodeon": "Melodeon",
                                "melodica": "Melodica", "mizmar": "Mizmar", "morin khuur": "Morin Khuur",
                                "mouth harp": "Mouth Harp/Jews Harp",
                                "jews harp": "Mouth Harp/Jews Harp", "mouth": "Mouth/Beat Box",
                                "beat box": "Mouth/Beat Box",
                                "mridangam": "Mridangam", "mukkuri": "Mukkuri/Tonkori", "tonkori": "Mukkuri/Tonkori",
                                "musette": "Musette", "music box": "Music Box", "musical saw": "Musical Saw",
                                "ney": "Ney",
                                "ngoni": "Ngoni", "non-specific": "Non-specific", "novachord": "Novachord",
                                "oboe": "Oboe",
                                "ocarina": "Ocarina", "omnichord": "Omnichord", "orchestra": "Orchestra",
                                "organ": "Organ",
                                "wurlitzer": "Organ - Wurlitzer", "oud": "Oud", "pads": "Pads",
                                "palm court": "Palm Court/Salon Orchestra",
                                "pan pipes": "Pan Pipes", "percussion": "Percussion", "piano": "Piano", "pipa": "Pipa",
                                "pipes": "Pipes", "pipes - celtic": "Pipes - Celtic",
                                "pipes - hornpipe": "Pipes - Hornpipe",
                                "pipes - pan": "Pipes - Pan", "polyphone": "Polyphone", "quena": "Quena",
                                "rabbi": "Rabbi",  # Assuming 'rabbi' is correct, not a typo for rabab
                                "rackett": "Rackett", "rainstick": "Rainstick", "ranat": "Ranat", "ratchet": "Ratchet",
                                "rebec": "Rebec", "recorder": "Recorder", "reed aerophone": "Reed Aerophone",
                                "riq": "Riq/Kanjira",
                                "kanjira": "Riq/Kanjira", "rubab": "Rubab/Robab/Rabab", "robab": "Rubab/Robab/Rabab",
                                "rabab": "Rubab/Robab/Rabab", "sackbut": "Sackbut", "sanshin": "Sanshin",
                                "santoor": "Santoor",
                                "sarangi": "Sarangi", "sarod": "Sarod", "saunter": "Saunter",
                                # "Saunter" as an instrument?
                                "saxophone": "Saxophone", "saz lute": "Saz Lute/Baglama", "baglama": "Saz Lute/Baglama",
                                "scheitholt": "Scheitholt", "scratching": "Scratching", "sfx": "SFX (Sound Effects)",
                                "effects": "SFX (Sound Effects)", "shaker": "Shaker", "shakuhachi": "Shakuhachi",
                                "shamisen": "Shamisen", "shawm": "Shawm", "shekere": "Shekere", "shenai": "Shenai",
                                "sho": "Sho",
                                "shou": "Shou", "side drum": "Side Drum", "singing bowls": "Singing Bowls",
                                "sitar": "Sitar",
                                "snare drum": "Snare Drum", "sound design": "Sound Design", "spoons": "Spoons",
                                "steel drums": "Steel Drums", "steelpan": "Steelpan", "sticks": "Sticks",
                                "string ensemble": "String Ensemble", "string quartet": "String Quartet",
                                "string section": "String Section",
                                "strings": "Strings", "suling": "Suling", "suona": "Suona", "surbahar": "Surbahar",
                                "surdo": "Surdo", "synthesizer": "Synthesizer", "synth": "Synthesizer",
                                "synths": "Synthesizer",
                                "tabla": "Tabla",
                                "taiko drum": "Taiko Drum", "talking drum": "Talking Drum", "tambourine": "Tambourine",
                                "tambura": "Tambura", "tar": "Tar", "tarabuka": "Tarabuka",
                                "temple bell": "Temple Bell",
                                "temple blocks": "Temple Blocks", "theremin": "Theremin",
                                "thunder sheet": "Thunder Sheet",
                                "tibetan singing bowls": "Tibetan Singing Bowls", "timbale": "Timbale",
                                "timpani": "Timpani",
                                "tiompan": "Tiompan", "tom toms": "Tom Toms", "toms": "Tom Toms",
                                "tongue drum": "Tongue Drum",
                                "toy instruments": "Toy Instruments", "transverse flute": "Transverse Flute",
                                "trautonium": "Trautonium",
                                "triangle": "Triangle", "tromba marina": "Tromba Marina", "trombone": "Trombone",
                                "trumpet": "Trumpet", "tuba": "Tuba", "udu": "Udu", "ukulele": "Ukulele",
                                "vibraphone": "Vibraphone", "vibraslap": "Vibraslap", "viol": "Viol", "viola": "Viola",
                                "viola da gamba": "Viola Da Gamba", "violin": "Violin", "vox": "Vocals", "vocal": "Vocals", "vocals": "Vocals",
                                "vocoder": "Vocoder",
                                "washboard": "Washboard", "waterphone": "Waterphone", "whip": "Whip",
                                "whisper": "Whisper",
                                "whistle": "Whistle", "wind chimes": "Wind Chimes", "wood block": "Wood Block",
                                "woodblock": "Wood Block", "woodwinds": "Woodwinds", "xiao": "Xiao",
                                "xylophone": "Xylophone",
                                "yangqin": "Yangqin", "zagat": "Zagat", "zither": "Zither",
                                "zourna/sorna/zurna": "Zourna/Sorna/Zurna",
                                "sorna": "Zourna/Sorna/Zurna", "zurna": "Zourna/Sorna/Zurna"
                            }
                            SORTED_INSTRUMENT_KEYWORDS = sorted(INSTRUMENT_KEYWORD_MAP.keys(), key=len, reverse=True)
                            val_for_Y = ""
                            fmt_stem_lower = fmt_stem.lower() if fmt_stem else ""
                            if fmt_stem_lower:
                                for keyword in SORTED_INSTRUMENT_KEYWORDS:
                                    if keyword == "percussion":
                                        if keyword in fmt_stem_lower:
                                            val_for_Y = INSTRUMENT_KEYWORD_MAP[keyword]
                                            break
                                    elif re.search(r'\b' + re.escape(keyword) + r'\b', fmt_stem_lower):
                                        val_for_Y = INSTRUMENT_KEYWORD_MAP[keyword]
                                        break
                            if val_for_Y:
                                df_processed.iloc[row_idx, Y_IDX] = val_for_Y
                        
                        # Populate other columns for ALL stem rows
                        if K_IDX < current_df_shape[1]: 
                            df_processed.iloc[row_idx, K_IDX] = os.path.splitext(str(fn_b))[0]
                        
                        if C_IDX < current_df_shape[1]:
                            p_val = str(df_processed.iloc[row_idx, P_IDX]) if P_IDX < current_df_shape[1] and pd.notna(df_processed.iloc[row_idx, P_IDX]) else ""
                            df_processed.iloc[row_idx, C_IDX] = f"{p_val} {main_tt_current_row} STEM {fmt_stem}".strip()
                        
                        if E_IDX < current_df_shape[1]:
                            df_processed.iloc[row_idx, E_IDX] = get_col_E_value_from_filename(str(fn_b))

                        if S_IDX < current_df_shape[1]: 
                            df_processed.iloc[row_idx, S_IDX] = f"STEM {fmt_stem}".strip()
                        
                        if T_IDX < current_df_shape[1]:
                            if is_vocal:
                                val_T = "Submix, Song, Lyrics, Vocals"
                                if match_src_row_for_generic_copy is not None and T_IDX < len(match_src_row_for_generic_copy):
                                    source_T_val_original = str(match_src_row_for_generic_copy.iloc[T_IDX])
                                    modified_T_val = re.sub(r'\bFull\b', 'Submix', source_T_val_original, flags=re.IGNORECASE, count=1)
                                    if modified_T_val != source_T_val_original:
                                        val_T = modified_T_val
                                    elif source_T_val_original.strip():
                                        val_T = source_T_val_original
                                df_processed.iloc[row_idx, T_IDX] = val_T
                            else:
                                df_processed.iloc[row_idx, T_IDX] = "Submix, No Lyrics, No Vocals"
                        
                        if U_IDX < current_df_shape[1]: 
                            df_processed.iloc[row_idx, U_IDX] = "N"

                        if AI_IDX < current_df_shape[1]:
                            if is_vocal and match_src_row_for_generic_copy is not None and AI_IDX < len(match_src_row_for_generic_copy):
                                df_processed.iloc[row_idx, AI_IDX] = match_src_row_for_generic_copy.iloc[AI_IDX]

                        if BC_IDX < current_df_shape[1]:
                            val_bc = "1" if is_vocal else "0"
                            df_processed.iloc[row_idx, BC_IDX] = val_bc
                            if BD_IDX < current_df_shape[1]:
                                if val_bc == "1":
                                    if fmt_stem_lower == "vocal background" or fmt_stem_lower == "vocals background":
                                        df_processed.iloc[row_idx, BD_IDX] = "Vocal Textures - Vocal Background"
                                    elif match_src_row_for_generic_copy is not None and BD_IDX < len(match_src_row_for_generic_copy):
                                        df_processed.iloc[row_idx, BD_IDX] = match_src_row_for_generic_copy.iloc[BD_IDX]
                                elif val_bc == "0":
                                    df_processed.iloc[row_idx, BD_IDX] = "No Vocal"

                    if file_was_modified:
                        output_buffer = io.BytesIO()
                        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                            df_processed.to_excel(writer, index=False, sheet_name='Sheet1')
                            worksheet = writer.sheets['Sheet1']
                            auto_adjust_column_width(worksheet)
                        processed_data_outputs.append((f"{uploaded_file_obj.name}", output_buffer.getvalue()))
                        processed_files_count += 1
                except IndexError as e_idx:
                    idx_arg = e_idx.args[0] if e_idx.args else -1
                    col_letter_involved = index_to_excel_col(idx_arg if isinstance(idx_arg, int) else -1)
                    current_file_status.error(
                        f"Index Error in {uploaded_file_obj.name}: {e_idx}. Problem with col {col_letter_involved}. File has {current_df_shape[1]} cols.");
                    skipped_files_count += 1
                except Exception as e:
                    current_file_status.error(f"Error processing {uploaded_file_obj.name}: {e}");
                    import traceback;

                    st.error(traceback.format_exc())
                    skipped_files_count += 1
                overall_progress_bar.progress((i + 1) / len(uploaded_files))

            current_file_status.empty()
            if processed_files_count == 0 and skipped_files_count == 0 and len(uploaded_files) > 0:
                st.info("Processing complete. No files were modified or met criteria for changes.")
            elif processed_files_count > 0 or skipped_files_count > 0:
                st.success(
                    f"Batch processing complete! {processed_files_count} file(s) processed, {skipped_files_count} file(s) skipped/errored.")

        with download_trigger_area:
            if not processed_data_outputs:
                if len(uploaded_files) > 0:
                    st.info("No files were processed that require downloading.")
            elif len(processed_data_outputs) == 1:
                st.success("One file processed. Download should start automatically...")
                fname, data_val = processed_data_outputs[0]
                html_dl = trigger_download_component(data_val, fname,
                                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.components.v1.html(html_dl, height=0, scrolling=False)
            elif len(processed_data_outputs) == 2:
                st.success("Two files processed. Downloads should start automatically (staggered)...")
                for idx, (fname, data_val) in enumerate(processed_data_outputs):
                    if idx > 0: time.sleep(1)
                    html_dl = trigger_download_component(data_val, fname,
                                                         "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.components.v1.html(html_dl, height=0, scrolling=False)
            elif len(processed_data_outputs) > 2:
                st.success(
                    f"{len(processed_data_outputs)} files processed. Zipping and download should start automatically...")
                zip_buffer = io.BytesIO();
                zip_filename = f"processed_files_{uuid.uuid4().hex[:8]}.zip"
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for fname, data_val in processed_data_outputs: zf.writestr(fname, data_val)
                zip_data = zip_buffer.getvalue()
                html_dl = trigger_download_component(zip_data, zip_filename, "application/zip")
                st.components.v1.html(html_dl, height=0, scrolling=False)
            if processed_data_outputs:
                st.caption("If downloads don't start, check browser pop-up/download settings.")
