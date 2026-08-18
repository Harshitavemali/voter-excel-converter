import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
    page_title="Telugu Voter Text to Excel Converter",
    page_icon="📄",
    layout="wide"
)

st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background-color: #0066cc;
        color: white;
        font-weight: bold;
        padding: 0.5rem 1rem;
        border-radius: 0.5rem;
    }
    .stButton>button:hover {
        background-color: #0052a3;
    }
    h1 {
        color: #0066cc;
    }
    .subtitle {
        color: #666;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.header("📖 How to Use")
    st.markdown("""
    ### Steps:
    1. **Copy Text from PDF**: Open your PDF and copy all the text (Ctrl+A, Ctrl+C)
    2. **Paste Text**: Paste the copied text into the text area
    3. **Process Text**: Click the 'Process Text' button to extract data
    4. **Preview Data**: Review the extracted voter information in the table
    5. **Download Excel**: Click the download button to get your Excel file
    
    ### Extracted Data Fields:
    - **S.No**: Serial Number
    - **Voter ID**: Voter identification number
    - **Voter Name**: Name of the voter
    - **Father/Husband Name**: Parent or spouse name
    - **House No**: Residence number
    - **Age**: Voter's age
    - **Gender**: Male (M) or Female (F)
    """)

TELUGU_NOISE_PATTERNS = [
    r'^Photo$', r'^Available$', r'^Available Available$',
    r'^వయస్సు లెక్కింపు తేదీ', r'^ప్రచురణ తేదీ',
    r'^శాసనసభ నియోజకవర్గం సంఖ్య మరియు పేరు', r'^విభాగం సంఖ్య మరియు పేరు',
    r'^భాగం సంఖ్య', r'^\d{2}-\d{2}-\d{4}$', r'^\|$',
]
TELUGU_NOISE_RE = re.compile('|'.join(TELUGU_NOISE_PATTERNS))

# "మొత్తం పేజీలు ... పేజీ N" marks a genuine page boundary in the source PDF.
# We use these to verify records/IDs counts stay balanced per page, so a
# mismatch or a merged/skipped record on one page can't silently cascade
# and shift every serial number on every page after it (as it would with a
# single document-wide positional match).
TELUGU_PAGE_BOUNDARY_RE = re.compile(r'^మొత్తం పేజీలు')

TELUGU_ID_RE = re.compile(r'([A-Z]{2,4})\s?(\d{6,7})')

# Matches one full record: name ... [relation-label ... relation-name] ... house ... age ... gender-word
# The relation group is OPTIONAL: the source PDF occasionally drops that field
# entirely for a record. Without this, such a record fails to match on its own
# and the regex swallows it into the NEXT record's match instead - merging two
# records into one and throwing off every serial number after it on that page.
# The name group also stops at the next "ఓటరు పేరు" marker (a second name
# occasionally appears before the first record's own details, another
# page-break artifact) - without this boundary, a record with no immediately
# following field would swallow the next record's name text too.
# Separators use [:|\s]* (not a fixed single ':') because OCR/copy sometimes doubles punctuation.
TELUGU_RECORD_RE = re.compile(
    r'ఓటరు\s*పేరు\s*[:|\s]*((?:(?!ఓటరు\s*పేరు).)+?)\s*'
    r'(?:(తండ్రి\s*పేరు|భర్త\s*పేరు)\s*[:|\s]*((?:(?!ఓటరు\s*పేరు).)+?)\s*)?'
    r'ఇంటి\s*సంఖ్య\s*[:|\s]*(.+?)\s*'
    r'వయస్సు\s*[:|\s]*(\d{1,6})\s*'
    r'లింగము\s*[:|\s]*(పురుషులు|స్త్రీలు)'
)

TELUGU_RELATION_LABEL_MAP = {'తండ్రి పేరు': 'F', 'భర్త పేరు': 'H'}
TELUGU_GENDER_MAP = {'పురుషులు': 'M', 'స్త్రీలు': 'F'}

def clean_text_telugu(text):
    lines = [l.strip() for l in text.split('\n')]
    out = []
    for l in lines:
        if not l:
            continue
        if TELUGU_PAGE_BOUNDARY_RE.match(l):
            out.append('\x00PAGE_BOUNDARY\x00')
            continue
        if TELUGU_NOISE_RE.match(l):
            continue
        out.append(l)
    return out

def split_into_telugu_pages(lines):
    pages = []
    current = []
    for l in lines:
        if l == '\x00PAGE_BOUNDARY\x00':
            pages.append(current)
            current = []
            continue
        current.append(l)
    if current:
        pages.append(current)
    return pages

def _dedupe_telugu_age(raw):
    # OCR/copy sometimes doubles the age digits, e.g. "4545" -> "45", "7474" -> "74"
    if len(raw) >= 4 and len(raw) % 2 == 0:
        half = len(raw) // 2
        if raw[:half] == raw[half:]:
            return raw[:half]
    if len(raw) > 3:
        return raw[:2]  # ages are virtually always 1-2 digits
    return raw

_TELUGU_STRAY_ID_RE = re.compile(r'[A-Z]{2,4}\s?\d{6,7}')
_TELUGU_STRAY_NOISE_RE = re.compile(r'\b(Photo|Available)\b', re.IGNORECASE)
_TELUGU_STRAY_NUMBER_RE = re.compile(r'(?<!\d)\d{1,4}(?!\d)')

def _clean_telugu_name_field(raw):
    # Strip stray voter IDs / serial numbers / "Photo"/"Available" that sometimes
    # get copy-pasted in between a record's own fields (mixed-column OCR artifact).
    cleaned = _TELUGU_STRAY_ID_RE.sub('', raw)
    cleaned = _TELUGU_STRAY_NOISE_RE.sub('', cleaned)
    cleaned = _TELUGU_STRAY_NUMBER_RE.sub('', cleaned)
    return re.sub(r'\s+', ' ', cleaned).strip(' |')

def _extract_telugu_serials(page_lines):
    """Extract printed serial numbers from a page's lines, in appearance order.
    A line counts as serial-bearing only if, after removing any Voter ID tokens,
    everything left on it is purely one or more short digit groups (no Telugu
    text, no age/house labels) - e.g. '1', '4 ATM0261701', or 'ATM2065589 11
    ATM1559012 12 ATM0261644'. This deliberately excludes age values (which sit
    on lines with 'వయస్సు') and house numbers (which contain hyphens/slashes).
    """
    serials = []
    for line in page_lines:
        stripped = TELUGU_ID_RE.sub(' ', line).strip()
        if not stripped:
            continue
        if re.fullmatch(r'(\d{1,4}\s*)+', stripped):
            serials.extend(re.findall(r'\d{1,4}', stripped))
    return [int(s) for s in serials]

def parse_voter_data_telugu(text):
    lines = clean_text_telugu(text)
    pages = split_into_telugu_pages(lines)

    voters = []
    seen_ids = set()
    last_known_serial = 0
    for page_lines in pages:
        page_joined = re.sub(r'\s+', ' ', ' '.join(page_lines))
        ids = [m.group(1) + m.group(2) for m in TELUGU_ID_RE.finditer(page_joined)]
        records = list(TELUGU_RECORD_RE.finditer(page_joined))
        printed_serials = _extract_telugu_serials(page_lines)
        page_count_mismatch = (len(records) != len(ids))
        serials_count_mismatch = (len(records) != len(printed_serials))

        for idx, m in enumerate(records):
            serial_inferred = False
            if idx < len(printed_serials) and not serials_count_mismatch:
                sno = printed_serials[idx]
                last_known_serial = sno
            else:
                sno = last_known_serial + 1
                last_known_serial = sno
                serial_inferred = True

            name_raw = re.sub(r'\s+', ' ', m.group(1)).strip(' |')
            name = _clean_telugu_name_field(name_raw)
            relation_missing = m.group(2) is None
            relation_type = TELUGU_RELATION_LABEL_MAP.get((m.group(2) or '').strip(), '')
            relation_raw = re.sub(r'\s+', ' ', m.group(3) or '').strip(' |')
            relation_name = _clean_telugu_name_field(relation_raw)
            house_raw = re.sub(r'\s+', ' ', m.group(4)).strip(' |')
            house_match = re.match(r'[\d][\d./\-]*', house_raw.lstrip())
            if not house_match:
                house_match = re.search(r'[\d][\d./\-]*', house_raw)
            house = house_match.group(0) if house_match else house_raw
            age = _dedupe_telugu_age(m.group(5))
            gender = TELUGU_GENDER_MAP.get(m.group(6), '')

            voter_id = ids[idx] if idx < len(ids) else ''
            needs_review = page_count_mismatch or serial_inferred or relation_missing
            if not voter_id:
                needs_review = True
            elif voter_id in seen_ids:
                needs_review = True
            else:
                seen_ids.add(voter_id)
            # If cleanup had to strip something, flag it so it's easy to spot-check.
            if name_raw != name or relation_raw != relation_name:
                needs_review = True

            voters.append({
                'S.No': sno,
                'Voter_ID': voter_id,
                'Voter_Name': name,
                'Father_Husband_Name': relation_name,
                'Relation_Type': relation_type,
                'House_No': house,
                'Age': age,
                'Gender': gender,
                'Needs_Review': 'Yes' if needs_review else ''
            })

    return voters

ENGLISH_NOISE_PATTERNS = [
    r'^Photo$', r'^photo$', r'^Available$', r'^PhotO$', r'^photO$',
    r'^Age as on', r'^Date of Publication', r'^Assembly Constituency No and Name',
    r'^Section No and Name', r'^Part No\.', r'^\d{2}-\d{2}-\d{4}$',
    r'^SUMMARY OF ELECTORS', r'^A\) NUMBER OF ELECTORS', r'^Roll Type',
    r'^Mother Roll$', r'^NUMBER OF ELECTORS$', r'^Signature of Electoral',
    r'^Draft Electoral Roll', r'^Male$', r'^Female$', r'^Third$', r'^Gender$',
    r'^Total$',
]
ENGLISH_NOISE_RE = re.compile('|'.join(ENGLISH_NOISE_PATTERNS), re.IGNORECASE)

# "Total Pages ... - Page N" marks a genuine page boundary in the source PDF
# (unlike the other noise lines above, which interject mid-column between a
# records-batch and its own matching IDs-batch and must NOT split them apart).
# We use these to verify records/IDs counts stay balanced per page, so a
# mismatch on one page is contained instead of silently shifting every
# record after it for the rest of the document.
ENGLISH_PAGE_BOUNDARY_RE = re.compile(r'^Total Pages', re.IGNORECASE)

ENGLISH_ID_LETTER_RE = re.compile(r'^([A-Z]{2,4})\s?(\d{6,7})$')
ENGLISH_ID_DIGIT_ONLY_RE = re.compile(r'^\d{6,7}$')

def clean_text_english_lines(text):
    lines = [l.strip() for l in text.split('\n')]
    out = []
    for l in lines:
        if not l:
            continue
        if ENGLISH_PAGE_BOUNDARY_RE.match(l):
            out.append('\x00PAGE_BOUNDARY\x00')  # keep as a marker, strip everything else about the line
            continue
        if ENGLISH_NOISE_RE.match(l):
            continue
        out.append(l)
    return out

def split_into_english_pages(lines):
    pages = []
    current = []
    for l in lines:
        if l == '\x00PAGE_BOUNDARY\x00':
            pages.append(current)
            current = []
            continue
        current.append(l)
    if current:
        pages.append(current)
    return pages

def _looks_like_stray_id(line):
    return bool(ENGLISH_ID_LETTER_RE.match(line) or ENGLISH_ID_DIGIT_ONLY_RE.match(line))

def _parse_one_english_record(lines, i, serial, inferred=False):
    n = len(lines)
    name_line = lines[i]
    name = name_line.split(':', 1)[1].strip() if ':' in name_line else ''
    i += 1
    stray_id_skipped = False
    # absorb wrapped continuation of name until relation label or House Number,
    # but skip (don't absorb) any stray Voter ID line that leaked in from a
    # page-break splitting the record's own fields apart.
    while i < n and not re.match(r'(Fathers Name|Husbands Name|Mothers Name|Others)', lines[i], re.I) \
            and not lines[i].lower().startswith('house number'):
        if _looks_like_stray_id(lines[i]):
            stray_id_skipped = True
            i += 1
            continue
        name += ' ' + lines[i]
        i += 1

    relation = ''
    relation_type = ''
    if i < n and re.match(r'(Fathers Name|Husbands Name|Mothers Name|Others)', lines[i], re.I):
        m = re.match(r'(Fathers Name|Husbands Name|Mothers Name|Others)\s*:?\s*(.*)', lines[i], re.I)
        label = m.group(1).lower()
        relation_type = {'fathers name': 'F', 'husbands name': 'H', 'mothers name': 'M', 'others': 'O'}.get(label, '')
        relation = m.group(2).strip()
        i += 1
        while i < n and not lines[i].lower().startswith('house number'):
            if _looks_like_stray_id(lines[i]):
                stray_id_skipped = True
                i += 1
                continue
            relation += ' ' + lines[i]
            i += 1

    house = ''
    if i < n and lines[i].lower().startswith('house number'):
        house = lines[i].split(':', 1)[1].strip() if ':' in lines[i] else ''
        i += 1

    age = ''
    gender = ''
    if i < n and lines[i].lower().startswith('age'):
        m = re.match(r'Age\s*:\s*(\d+)\s*Gender\s*:\s*(\w+)', lines[i], re.I)
        if m:
            age = m.group(1)
            gender_raw = m.group(2)
            gender = {'male': 'M', 'female': 'F', 'm': 'M', 'f': 'F'}.get(gender_raw.lower(), gender_raw)
        i += 1

    rec = {
        'S.No': serial,
        'Voter_Name': re.sub(r'\s+', ' ', name).strip(),
        'Father_Husband_Name': re.sub(r'\s+', ' ', relation).strip(),
        'Relation_Type': relation_type,
        'House_No': house,
        'Age': age,
        'Gender': gender,
        'Serial_Inferred': inferred,
        'Stray_ID_Skipped': stray_id_skipped
    }
    return rec, i


def _extract_english_records(lines):
    # Records come out grouped by column (e.g. serials 1,4,7,10...), NOT interleaved with IDs.
    records = []
    i = 0
    n = len(lines)
    last_serial = None  # used to infer a serial dropped by OCR (+3 pattern within a column)
    while i < n:
        line = lines[i]
        if re.fullmatch(r'\d{1,4}', line) and i + 1 < n and lines[i + 1].lower().startswith('name'):
            serial = line
            i += 1
            rec, i = _parse_one_english_record(lines, i, serial, inferred=False)
            records.append(rec)
            last_serial = int(serial)
            continue
        elif line.lower().startswith('name'):
            inferred_serial = str(last_serial + 3) if last_serial is not None else ''
            rec, i = _parse_one_english_record(lines, i, inferred_serial, inferred=True)
            records.append(rec)
            if inferred_serial:
                last_serial = int(inferred_serial)
            continue
        i += 1
    return records


def _extract_english_ids(lines):
    # IDs come out in a separate batch right after each column's records, in the same order.
    ids = []
    for line in lines:
        m = ENGLISH_ID_LETTER_RE.match(line)
        if m:
            ids.append(m.group(1) + m.group(2))
            continue
        if ENGLISH_ID_DIGIT_ONLY_RE.match(line):
            ids.append(line)  # letter prefix dropped by OCR/extraction
    return ids


def parse_voter_data_english(text):
    lines = clean_text_english_lines(text)
    pages = split_into_english_pages(lines)

    voters = []
    seen_ids = set()
    for page_lines in pages:
        records = _extract_english_records(page_lines)
        ids = _extract_english_ids(page_lines)
        # If this page's records/IDs counts don't match, something on THIS page
        # (a dropped record, a stray/missing ID) desynced the pairing - flag
        # every record on this page rather than let a wrong guess ride, and
        # contain the problem here instead of letting it shift every record
        # on every later page too.
        page_count_mismatch = (len(records) != len(ids))

        for idx, rec in enumerate(records):
            voter_id = ids[idx] if idx < len(ids) else ''
            needs_review = page_count_mismatch
            if not voter_id:
                needs_review = True
            elif voter_id in seen_ids:
                needs_review = True
            else:
                seen_ids.add(voter_id)
            if rec.get('Serial_Inferred'):
                needs_review = True
            if rec.get('Stray_ID_Skipped'):
                needs_review = True

            voters.append({
                'S.No': rec['S.No'],
                'Voter_ID': voter_id,
                'Voter_Name': rec['Voter_Name'],
                'Father_Husband_Name': rec['Father_Husband_Name'],
                'Relation_Type': rec.get('Relation_Type', ''),
                'House_No': rec['House_No'],
                'Age': rec['Age'],
                'Gender': rec['Gender'],
                'Needs_Review': 'Yes' if needs_review else ''
            })

    return voters

def parse_voter_data(text, language):
    if language == 'Telugu':
        return parse_voter_data_telugu(text)
    else:
        return parse_voter_data_english(text)

def convert_to_excel(data):
    df = pd.DataFrame(data)
    df = df[['S.No', 'Voter_ID', 'Voter_Name', 'Relation_Type', 'Father_Husband_Name', 'House_No', 'Age', 'Gender', 'Needs_Review']]
    df['_sort_key'] = pd.to_numeric(df['S.No'], errors='coerce')
    df = df.sort_values('_sort_key', na_position='last').drop(columns='_sort_key').reset_index(drop=True)
    df.columns = ['S.No', 'Voter ID', 'Voter Name', 'H/F', 'Father/Husband Name', 'House No', 'Age', 'Gender', 'Needs Review']
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Voter Data')
        
        workbook = writer.book
        worksheet = writer.sheets['Voter Data']
        
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#0066cc',
            'font_color': 'white',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        worksheet.set_column('A:A', 8)
        worksheet.set_column('B:B', 15)
        worksheet.set_column('C:C', 30)
        worksheet.set_column('D:D', 8)
        worksheet.set_column('E:E', 30)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:G', 8)
        worksheet.set_column('H:H', 10)
        worksheet.set_column('I:I', 12)
    
    output.seek(0)
    return output

st.title("📄 Voter Text to Excel Converter")
st.markdown('<p class="subtitle">Paste copied text from a voter roll PDF (English or Telugu) and convert to Excel format</p>', unsafe_allow_html=True)

st.markdown("---")

st.markdown("**Select PDF Language:**")
language = st.radio(
    "Select PDF Language:",
    options=["English", "Telugu"],
    horizontal=True,
    label_visibility="collapsed"
)

pasted_text = st.text_area(
    f"Paste the copied text from your {language} PDF here:",
    height=300,
    placeholder="Copy all text from your PDF (Ctrl+A, Ctrl+C) and paste it here (Ctrl+V)..."
)

st.markdown("---")

if pasted_text and len(pasted_text.strip()) > 0:
    st.success(f"✅ Text pasted successfully: **{len(pasted_text):,}** characters")
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        process_button = st.button("🔄 Process Text", use_container_width=True)
    
    if process_button:
        with st.spinner(f"⏳ Processing {language} voter data from pasted text..."):
            voter_data = parse_voter_data(pasted_text, language)
            
            if voter_data:
                st.session_state['voter_data'] = voter_data
                st.session_state['processed'] = True
                st.success(f"✅ Successfully extracted **{len(voter_data)}** voter records!")
            else:
                st.error("❌ No voter data found in the pasted text. Make sure you copied the text correctly from the PDF.")

if 'processed' in st.session_state and st.session_state['processed']:
    voter_data = st.session_state['voter_data']
    
    st.markdown("---")
    st.subheader(f"📊 Extracted Voter Data ({len(voter_data)} records)")
    
    search_term = st.text_input("🔍 Search by Voter ID or Name:", "")
    
    df = pd.DataFrame(voter_data)
    df = df[['S.No', 'Voter_ID', 'Voter_Name', 'Relation_Type', 'Father_Husband_Name', 'House_No', 'Age', 'Gender', 'Needs_Review']]
    df['_sort_key'] = pd.to_numeric(df['S.No'], errors='coerce')
    df = df.sort_values('_sort_key', na_position='last').drop(columns='_sort_key').reset_index(drop=True)
    df.columns = ['S.No', 'Voter ID', 'Voter Name', 'H/F', 'Father/Husband Name', 'House No', 'Age', 'Gender', 'Needs Review']
    
    if search_term:
        mask = df['Voter ID'].str.contains(search_term, case=False, na=False) | \
               df['Voter Name'].str.contains(search_term, case=False, na=False)
        df_display = df[mask]
        st.write(f"Showing {len(df_display)} matching records")
    else:
        df_display = df
    
    st.dataframe(df_display, use_container_width=True, height=400)
    
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        excel_file = convert_to_excel(voter_data)
        st.download_button(
            label="⬇️ Download Excel File",
            data=excel_file,
            file_name=f"voter_data_{len(voter_data)}_records.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    st.markdown("---")
    st.subheader("📈 Data Statistics")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Voters", len(voter_data))
    male_count = sum(1 for v in voter_data if v.get('Gender', '') == 'M')
    col2.metric("Male (M)", male_count)
    female_count = sum(1 for v in voter_data if v.get('Gender', '') == 'F')
    col3.metric("Female (F)", female_count)
    col4.metric("Gender Unknown", len(voter_data) - male_count - female_count)