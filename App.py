import pandas as pd
import re
from difflib import SequenceMatcher

# ===== CONFIG =====
main_file = "main.xlsx"
output_file = "output.xlsx"

# ===== STEP 1: Choose Sheet =====
sheet_names = pd.ExcelFile(main_file).sheet_names
print("Available sheets:", sheet_names)
sheet_name = input("Enter the sheet name to process: ").strip()

# ===== STEP 2: Detect Real Header Row =====
temp_df = pd.read_excel(main_file, sheet_name=sheet_name, header=None)
header_row_idx = None
for i, row in temp_df.iterrows():
    if row.astype(str).str.strip().str.upper().eq("DESCRIPTION").any():
        header_row_idx = i
        break

if header_row_idx is None:
    raise ValueError("❌ Could not find a column named 'DESCRIPTION' in this sheet.")

df = pd.read_excel(main_file, sheet_name=sheet_name, header=header_row_idx)

# ===== STEP 3: Ask for Description Column =====
print("Detected columns:", list(df.columns))
desc_col = input("Enter the column name that contains the description: ").strip()

# ===== STEP 4: Load Mapping =====
mapping_df = pd.read_excel(main_file, sheet_name="BizGroup2024")
mapping_df.columns = mapping_df.columns.str.strip()
mapping_df["PRODUCT LINE"] = mapping_df["PRODUCT LINE"].astype(str).str.strip()
mapping_df["BUSINESS GROUP"] = mapping_df["BUSINESS GROUP"].astype(str).str.strip()

# ===== STEP 5: Extraction Functions =====
def extract_npi_name(text):
    if pd.isna(text):
        return "N/A"
    text = str(text)
    
    # Pattern 1: Original "NPI Project Name:" format
    try:
        start = text.index("NPI Project Name:") + 17
        end = text.index("Product Model", start)
        return text[start:end].strip()
    except ValueError:
        pass
    
    # Pattern 2: "NPI---" format
    try:
        if "NPI---" in text:
            start = text.index("NPI---") + 6
            # Look for common endings like comma, period, or specific phrases
            possible_ends = [",", ".", "Rev", "BOM", "Project Manager", "FYI"]
            end_positions = []
            for end_marker in possible_ends:
                pos = text.find(end_marker, start)
                if pos != -1:
                    end_positions.append(pos)
            if end_positions:
                end = min(end_positions)  # Take the earliest ending
                return text[start:end].strip()
    except ValueError:
        pass
    
    # Pattern 3: formats like "NPI - Name ...", "NPI, Name ..." or "NPI-Name (parts...)"
    try:
        # Accept comma, colon, dash, or whitespace after NPI
        m = re.search(
            r"\bNPI\b\s*[-:,\s]*\s*(.+?)(?:\(|\bBOM\b|\bProject Manager\b|\bFYI\b|\bProduct Line\b|$|\n)",
            text,
            flags=re.IGNORECASE,
        )
        if m:
            name = m.group(1).strip()
            # Remove common trailing annotations like 'BOM load only', 'Complete data package release', etc.
            name = re.sub(r"\s*BOM.*$", "", name, flags=re.IGNORECASE)
            name = re.sub(r"\s*Complete data package.*$", "", name, flags=re.IGNORECASE)
            name = name.rstrip(' :,-?')
            # Collapse multiple spaces
            name = re.sub(r"\s+", " ", name)
            if name:
                return name
    except Exception:
        pass
        
    return "N/A"

def extract_model_number(text):
    if pd.isna(text):
        return "N/A"
    text = str(text)
    
    # Pattern 1: Original "Product Model:" format
    try:
        start = text.index("Product Model:") + 14
        end = text.index("\n", start) if "\n" in text[start:] else len(text)
        return text[start:end].strip()
    except ValueError:
        pass
    
    # Pattern 2: "Model No.;" format
    try:
        if "Model No.;" in text:
            start = text.index("Model No.;") + 10
            # Look for common endings like "Effectivity Date", "Change Type", etc.
            possible_ends = ["Effectivity Date", "Change Type", "Effective Date", "\n"]
            end_positions = []
            for end_marker in possible_ends:
                pos = text.find(end_marker, start)
                if pos != -1:
                    end_positions.append(pos)
            if end_positions:
                end = min(end_positions)
                models = text[start:end].strip()
                # Join all models with commas
                return ", ".join(model.strip() for model in models.split(','))
    except ValueError:
        pass
    
    # Pattern 3: Direct part number mention with common Keysight/Agilent patterns
    # Looking for patterns like XXXXX-XXXXX or XXXX-XXXX where X is digit
    part_number_patterns = [
        r'\b\d{5}-\d{5}\b',  # e.g., 54955-66401
        r'\b\d{4}-\d{4}\b',  # e.g., 1827-8247
        r'\b[A-Z]\d{4}[A-Z]?\b'  # e.g., N5227B, E5055A
    ]
    
    all_matches = []
    for pattern in part_number_patterns:
        matches = re.findall(pattern, text)
        if matches:
            # Add all matches to the list
            for match in matches:
                if match not in all_matches:  # Avoid duplicates
                    all_matches.append(match)
    
    if all_matches:
        return ", ".join(all_matches)
    
    return "N/A"

def _strong_model_match(query, ref):
    """Check if there's a strong match between the query model and reference model."""
    if pd.isna(query) or pd.isna(ref):
        return False
    query = str(query).strip().upper()
    ref = str(ref).strip().upper()
    
    # Split reference cell in case it contains multiple models
    # Handle cases with parentheses by extracting both full text and content within parentheses
    ref_models = []
    for part in ref.split(","):
        part = part.strip()
        ref_models.append(part)
        # Extract models from parentheses
        paren_matches = re.findall(r'\((.*?)\)', part)
        for paren_content in paren_matches:
            # Split content by slashes or dashes for multiple models
            sub_models = re.split(r'[/\-]', paren_content)
            ref_models.extend([m.strip() for m in sub_models])
        # Also add any models separated by slashes outside parentheses
        if '/' in part:
            ref_models.extend([m.strip() for m in part.split('/')])
    
    # Handle query with multiple models (comma, space, or dash separated)
    query_models = re.split(r'[,\s-]+', query)
    
    for query_model in query_models:
        query_model = query_model.strip()
        if not query_model:
            continue
            
        # 1. First try exact match
        if any(query_model == ref_model for ref_model in ref_models):
            return True
            
        # 2. Extract base numbers from query (e.g., 85059 from 85059-20067)
        query_base_match = re.match(r'(\d+)', query_model)
        if query_base_match:
            query_base = query_base_match.group(1)
            
            for ref_model in ref_models:
                # Check if the reference model contains this base number
                if query_base in ref_model:
                    # Additional validation to ensure it's a real match
                    ref_base_match = re.match(r'(\d+)', ref_model)
                    if ref_base_match and ref_base_match.group(1) == query_base:
                        return True
                
                # Handle wildcard cases (xxxx)
                if 'X' in ref_model:
                    ref_pattern = ref_model.replace('X', r'\d')
                    if re.match(ref_pattern, query_model):
                        return True
        
        # 3. Try model family match for letter-prefixed models (e.g., N5227B)
        query_family_match = re.match(r'([A-Z]+\d+)', query_model)
        if query_family_match:
            query_family = query_family_match.group(1)
            for ref_model in ref_models:
                ref_family_match = re.match(r'([A-Z]+\d+)', ref_model)
                if ref_family_match and ref_family_match.group(1) == query_family:
                    return True
    
    return False

def _pick_matching_model_from_cell(ref_cell, model_number):
    """Pick the most appropriate model from a cell that might contain multiple models."""
    if pd.isna(ref_cell) or pd.isna(model_number):
        return str(ref_cell)
    ref_models = [model.strip() for model in str(ref_cell).split(",")]
    model_number = str(model_number).strip()
    
    # If exact match exists, return it
    for ref_model in ref_models:
        if ref_model.upper() == model_number.upper():
            return ref_model
    
    # Otherwise return the first one (or full cell if no comma)
    return ref_models[0] if ref_models else str(ref_cell)

def extract_model_number(text):
    if pd.isna(text):
        return "N/A"
    text = str(text)
    # Primary logic: capture after "(Product Model/)?Part Number(s) Affected:" up to the next section header
    # Examples covered:
    # - "Product Model/Part Number Affected: AD1011A & AD1012A Product Line: ..."
    # - "Product Model/Part Number Affected: 1855-2870 Effective Date: ..."
    m = re.search(
        r"(?:Product\s*Model\s*/\s*)?Part\s*Number(?:s)?\s*Affect(?:ed)?\s*:?:?\s*(.*?)\s*(?=(?:Product\s*Line|Effective|Brief\s*Description|Material\s*Disposition|Reason|Approval|$|\n|\r))",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if m:
        val = m.group(1).strip()
        # Normalize conjunctions and separators
        val = re.sub(r"\s*(?:&|\band\b)\s*", ", ", val, flags=re.IGNORECASE)
        # Trim trailing punctuation/separators
        val = re.sub(r"[\s,;\-/]+$", "", val)
        if val:
            return val

    # Secondary logic (fallback): after 'Model affected:' up to the word 'CM'
    # Example: 'Model affected: N9355G-ATO-51774 , ... CM ...'
    m = re.search(r"Model\s*affected\s*:\s*(.*?)(?=\s*CM\b)", text, re.IGNORECASE | re.DOTALL)
    if m:
        secondary = m.group(1).strip()
        # Trim trailing separators/punctuation
        secondary = re.sub(r"[\s,;\-]+$", "", secondary)
        if secondary:
            return secondary

    # Tertiary logic: capture the first model-like token with digits after 'Model affected:' up to next header
    m = re.search(
        r"Model\s*affected\s*:\s*(.*?)(?=(?:Product\s*Line|Effective|Brief\s*Description|Material\s*Disposition|Reason|Approval|$|\n|\r))",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if m:
        seg = m.group(1)
        mm = re.search(r"\b[A-Za-z]*\d+[A-Za-z0-9\-_/]*\b", seg)
        if mm:
            token = mm.group(0).strip()
            if token:
                return token

    return "N/A"

def extract_product_line(text):
    if pd.isna(text):
        return None
    text = str(text)
    # Try multiple patterns:
    # 1) "Product Line: <value>" stopping before keywords like Effective/Model/end
    # 2) "PL: <value>" stopping before keywords like Model/Effective/end
    patterns = [
        r"Product\s*Line\s*:?\s*(.*?)\s*(Effective\b|Model\b|Product\s*Model\b|$|\n|\r)",
        r"\bPL\s*:?\s*(.*?)\s*(Model\b|Product\s*Model\b|Effective\b|$|\n|\r)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
        if match:
            pl_value = match.group(1).strip()
            # Trim trailing separators/punctuation
            pl_value = re.sub(r"[\s,;\-]+$", "", pl_value)
            # Normalize leading 'PL' prefix if included in the captured group
            if pl_value.upper().startswith("PL"):
                pl_value = pl_value[2:].strip()
            if pl_value:
                return pl_value
    return None

def _parse_model_affected_segment(text):
    if pd.isna(text):
        return None, None
    s = str(text)
    # Take segment after 'Model affected:' up to next header/section break
    m = re.search(
        r"Model\s*affected\s*:\s*(.*?)(?=(?:Product\s*Line|Effective|Brief\s*Description|Material\s*Disposition|Reason|Approval|NPI\s*Project\s*Name|$|\n|\r))",
        s,
        re.IGNORECASE | re.DOTALL,
    )
    if not m:
        return None, None
    seg = m.group(1)
    # Find the first model-like token (letters optional + digits + optional suffix/hyphenated)
    mm = re.search(r"\b[A-Za-z]*\d+[A-Za-z0-9\-_/]*\b", seg)
    if not mm:
        return None, None
    model_token = mm.group(0).strip()
    # NPI candidate is text before that token
    npi_candidate = seg[: mm.start()].strip()
    # Clean npi text: collapse spaces, strip trailing separators
    npi_candidate = re.sub(r"\s+", " ", npi_candidate)
    npi_candidate = re.sub(r"[\s,;:\-_/]+$", "", npi_candidate)
    if npi_candidate == "":
        npi_candidate = None
    return npi_candidate, model_token

df["NPI Name"] = df[desc_col].apply(extract_npi_name)
df["Model Number"] = df[desc_col].apply(extract_model_number)
df["Product Line"] = df[desc_col].apply(extract_product_line)

# ===== STEP 5.0b: Fallback parse from 'Model affected:' to backfill NPI/Model if missing =====
parsed = df[desc_col].apply(_parse_model_affected_segment)
df["_Parsed NPI"], df["_Parsed Model"] = zip(*parsed)
def _is_invalid_simple(value) -> bool:
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    s = str(value).strip().lower()
    if s == "":
        return True
    return s in {"n/a", "na", "n.a.", "not available", "none", "null", "-"}

mask_model_missing = df["Model Number"].apply(_is_invalid_simple) & df["_Parsed Model"].notna()
df.loc[mask_model_missing, "Model Number"] = df.loc[mask_model_missing, "_Parsed Model"]
mask_npi_missing = df["NPI Name"].apply(_is_invalid_simple) & df["_Parsed NPI"].notna()
df.loc[mask_npi_missing, "NPI Name"] = df.loc[mask_npi_missing, "_Parsed NPI"]
df.drop(columns=["_Parsed NPI", "_Parsed Model"], inplace=True)

# ===== STEP 5.1: Reference Data =====
ref_df = pd.read_excel(main_file, sheet_name="ReferenceModelNumber")
ref_df["NPI NAME"] = ref_df["NPI NAME"].astype(str).str.strip()
ref_df["MODEL NUMBER"] = ref_df["MODEL NUMBER"].astype(str).str.strip()
# Create a shuffled copy of ref_df to avoid bias based on row order
ref_df_shuffled = ref_df.sample(frac=1.0, random_state=42).reset_index(drop=True)

def similar(a, b):
    return SequenceMatcher(None, a, b).ratio()

def _normalize(text: str) -> str:
    if text is None:
        return ""
    # Keep only alphanumerics, lowercase; drop spaces/punctuation/underscores
    return re.sub(r"[^0-9a-z]+", "", str(text).lower())

def _tokens(text: str):
    if text is None:
        return []
    # Split on non-alphanumerics; drop empties; lowercase
    return [t for t in re.split(r"[^0-9a-zA-Z]+", str(text).strip().lower()) if t]

def _is_model_like_token(token: str) -> bool:
    token = str(token)
    return any(c.isalpha() for c in token) and any(c.isdigit() for c in token)

def _model_in_cell(cell_value, target_value) -> bool:
    if _is_invalid_extracted(target_value):
        return False
    cell_str = str(cell_value)
    target_str = str(target_value)

    # Tokenize both sides
    cell_tokens_raw = [t for t in _tokens(cell_str) if _is_model_like_token(t)]
    target_tokens_raw = [t for t in _tokens(target_str) if _is_model_like_token(t)]

    # Quick exact token match
    cell_tokens_norm = {_normalize(t) for t in cell_tokens_raw}
    for tgt in target_tokens_raw:
        tgt_norm = _normalize(tgt)
        if tgt_norm in cell_tokens_norm:
            return True

    # Smarter family-aware match using parsed tokens
    def parse_list(tokens):
        parsed = []
        for tok in tokens:
            p = _parse_model_token(tok)
            if p:
                parsed.append(p)  # (alpha, num_str, suffix)
        return parsed

    cell_parsed = parse_list(cell_tokens_raw)
    target_parsed = parse_list(target_tokens_raw)

    for ta, tn, ts in target_parsed:
        for ca, cn, cs in cell_parsed:
            # Same alpha and same numeric core → accept regardless of suffix differences
            if ta == ca and tn == cn:
                return True
            # Prefix-safe: same alpha; numeric cores share prefix; one has longer numeric part
            if ta == ca and (tn.startswith(cn) or cn.startswith(tn)):
                # Ensure at least 3-digit overlap to avoid weak matches
                if len(tn) >= 3 and len(cn) >= 3 and (tn[:3] == cn[:3]):
                    return True

    # Fallback to normalized full-string containment when all else fails
    norm_cell = _normalize(cell_str)
    norm_target = _normalize(target_str)
    return bool(norm_cell) and (norm_target in norm_cell or norm_cell in norm_target)

def _is_invalid_extracted(value) -> bool:
    # Treat common placeholders as invalid for matching
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    s = str(value).strip().lower()
    if s == "":
        return True
    return s in {"n/a", "na", "n.a.", "not available", "none", "null", "-"}

def _model_stem(token: str) -> str:
    # Extract leading letters+digits stem, e.g., 'e7515e' -> 'e7515', 'm9502a' -> 'm9502'
    token = str(token).lower()
    m = re.match(r"^([a-z]*\d+)", token)
    return m.group(1) if m else token

def _parse_model_token(token: str):
    # Split into alpha prefix, numeric core, and optional suffix
    # Examples: 'm1742a' -> ('m', '1742', 'a'); 'e7515ea1' -> ('e', '7515', 'ea1')
    token = str(token).lower()
    m = re.match(r"^([a-z]+)(\d+)([a-z0-9]+)?$", token)
    if not m:
        return None
    alpha = m.group(1)
    num_str = m.group(2)
    suffix = m.group(3) or ""
    return alpha, num_str, suffix

def _series_key(alpha: str, num_str: str, series_digits: int = 3) -> str:
    # Series key: alpha + first N digits of numeric part
    return f"{alpha}{num_str[:series_digits]}"

def _best_match_row(extracted_value, ref_df, ref_col, *, fuzzy_threshold: float = 0.82, prefer_token: bool = False):
    if _is_invalid_extracted(extracted_value):
        return None

    extracted_raw = str(extracted_value).strip()
    extracted_lower = extracted_raw.lower()
    extracted_norm = _normalize(extracted_raw)
    extracted_tokens = set(_tokens(extracted_raw))

    # For model numbers, prioritize rows by overlap in model-like tokens (contain both letters and digits)
    def is_model_like(token: str) -> bool:
        return any(c.isalpha() for c in token) and any(c.isdigit() for c in token)

    if prefer_token:
        extracted_model_tokens = {t for t in extracted_tokens if is_model_like(t)}
        if extracted_model_tokens:
            # 0) Exact token overlap of model-like tokens
            overlap_candidates = []
            for _, row in ref_df.iterrows():
                ref_val_raw = str(row[ref_col]).strip()
                ref_tokens = {t for t in _tokens(ref_val_raw) if is_model_like(t)}
                overlap = len(extracted_model_tokens & ref_tokens)
                if overlap > 0:
                    overlap_candidates.append((overlap, len(_normalize(ref_val_raw)), row))
            if overlap_candidates:
                overlap_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
                return overlap_candidates[0][2]

            # 0a) Series family proximity (e.g., 'm1740a, m1741a, m1749a' -> closest 'm1742a')
            extracted_models = [
                _parse_model_token(t) for t in extracted_model_tokens
            ]
            extracted_models = [p for p in extracted_models if p is not None]
            extracted_series_keys = { _series_key(alpha, num_str) for alpha, num_str, _ in extracted_models }
            if extracted_series_keys:
                series_candidates = []
                for _, row in ref_df.iterrows():
                    ref_val_raw = str(row[ref_col]).strip()
                    ref_tokens_list = [t for t in _tokens(ref_val_raw) if is_model_like(t)]
                    best_series_overlap = 0
                    best_distance = None
                    best_suffix_letters_only = 0
                    best_suffix_len = 0
                    matched = False
                    for rt in ref_tokens_list:
                        parsed = _parse_model_token(rt)
                        if not parsed:
                            continue
                        r_alpha, r_num_str, r_suffix = parsed
                        r_series = _series_key(r_alpha, r_num_str)
                        if r_series in extracted_series_keys:
                            matched = True
                            # Overlap count increases per matching token
                            best_series_overlap += 1
                            r_num = int(r_num_str)
                            # Distance to nearest extracted numeric part with same alpha
                            distances = []
                            for e_alpha, e_num_str, _ in extracted_models:
                                if e_alpha == r_alpha:
                                    distances.append(abs(int(e_num_str) - r_num))
                            d = min(distances) if distances else 10**9
                            if best_distance is None or d < best_distance:
                                best_distance = d
                                best_suffix_letters_only = 1 if (r_suffix.isalpha() and len(r_suffix) > 0) else 0
                                best_suffix_len = len(r_suffix)
                    if matched:
                        series_candidates.append(((best_series_overlap, -best_distance if best_distance is not None else 0, best_suffix_letters_only, -best_suffix_len), len(_normalize(ref_val_raw)), row))
                if series_candidates:
                    # Prefer more tokens in same series, then closer numeric core, then letters-only and shorter suffix, then longer ref
                    series_candidates.sort(key=lambda x: (x[0][0], x[0][1], x[0][2], x[0][3], x[1]), reverse=True)
                    return series_candidates[0][2]

            # 0b) Family stem prefix match (e.g., 'e7515-20168' should match 'e7515a/e/p/...')
            stems = { _model_stem(t) for t in extracted_model_tokens }
            prefix_candidates = []
            for _, row in ref_df.iterrows():
                ref_val_raw = str(row[ref_col]).strip()
                ref_tokens = [t for t in _tokens(ref_val_raw) if is_model_like(t)]
                best_suffix_score = None
                best_suffix_len = None
                matched_any = 0
                for rt in ref_tokens:
                    for stem in stems:
                        if rt.startswith(stem) and rt != stem:
                            matched_any = 1
                            suffix = rt[len(stem):]
                            suffix_letters_only = 1 if suffix.isalpha() else 0
                            suffix_len = len(suffix)
                            # Keep the best (prefer letters-only and shorter suffix)
                            candidate_score = (suffix_letters_only, -suffix_len)
                            if best_suffix_score is None or candidate_score > best_suffix_score:
                                best_suffix_score = candidate_score
                                best_suffix_len = suffix_len
                if matched_any:
                    # Tie-break with normalized ref length to stabilize
                    prefix_candidates.append((best_suffix_score if best_suffix_score is not None else (0, 0), len(_normalize(ref_val_raw)), row))
            if prefix_candidates:
                prefix_candidates.sort(key=lambda x: (x[0][0], x[0][1], x[1]), reverse=True)
                return prefix_candidates[0][2]

    # 1) Exact match (case-insensitive)
    exact_candidates = []
    for _, row in ref_df.iterrows():
        ref_val_raw = str(row[ref_col]).strip()
        if extracted_lower == ref_val_raw.lower():
            exact_candidates.append((row, ref_val_raw))
    if exact_candidates:
        # If multiple, prefer token containment (for model numbers), otherwise pick the one with longest ref value
        if prefer_token:
            token_pref = []
            for row, ref_val_raw in exact_candidates:
                ref_tokens = set(_tokens(ref_val_raw))
                token_match = extracted_lower in ref_tokens
                token_pref.append((int(token_match), len(_normalize(ref_val_raw)), row))
            token_pref.sort(key=lambda x: (x[0], x[1]), reverse=True)
            return token_pref[0][2]
        # No token preference; choose the one with the longest normalized ref
        exact_candidates.sort(key=lambda item: len(_normalize(item[1])), reverse=True)
        return exact_candidates[0][0]

    # 2) Substring containment (case-insensitive, punctuation/spaces ignored)
    substring_candidates = []
    for _, row in ref_df.iterrows():
        ref_val_raw = str(row[ref_col]).strip()
        ref_norm = _normalize(ref_val_raw)
        if not ref_norm:
            continue
        contains = extracted_norm in ref_norm or ref_norm in extracted_norm
        if contains:
            token_match = False
            if prefer_token:
                ref_tokens = set(_tokens(ref_val_raw))
                token_match = extracted_lower in ref_tokens
            # Estimate match length: smaller of the two norms when one contains the other
            match_len = len(extracted_norm) if extracted_norm in ref_norm else len(ref_norm)
            substring_candidates.append((int(token_match), match_len, len(ref_norm), row))
    if substring_candidates:
        # Prefer token containment (only meaningful when prefer_token=True), then longest match length, then longer ref
        substring_candidates.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
        return substring_candidates[0][3]

    # 3) Fuzzy match (on normalized text)
    fuzzy_candidates = []
    for _, row in ref_df.iterrows():
        ref_val_raw = str(row[ref_col]).strip()
        ref_norm = _normalize(ref_val_raw)
        if not ref_norm:
            continue
        score = similar(extracted_norm, ref_norm)
        if score >= fuzzy_threshold:
            token_match = False
            if prefer_token:
                ref_tokens = set(_tokens(ref_val_raw))
                token_match = extracted_lower in ref_tokens
            fuzzy_candidates.append((int(token_match), score, len(ref_norm), row))
    if fuzzy_candidates:
        # Prefer token containment (for model numbers), then higher similarity, then longer ref
        fuzzy_candidates.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
        return fuzzy_candidates[0][3]

    return None

# ===== Pass 1: Self-to-self match with new matching flow =====
def match_by_npi(npi_name):
    """Enhanced NPI matching with split comparison."""

    if _is_invalid_extracted(npi_name):
        return None, None
        
    # Special-case: if 'blackhawk' appears in the extracted NPI name, reference the first NPI in the reference sheet that contains 'Blackhawk'
    if re.search(r'\bblackhawk\b|\bblackhawk\d+\b', str(npi_name).lower()):
        blackhawk_row = ref_df_shuffled[ref_df_shuffled['NPI NAME'].str.lower().str.contains(r'\bblackhawk\b', regex=True, na=False)]
        if not blackhawk_row.empty:
            return blackhawk_row.iloc[0]['NPI NAME'], blackhawk_row.iloc[0]['MODEL NUMBER']

    # Special-case: if 'hedgehog' appears anywhere in the extracted NPI name, reference the first 'Hedgehog' NPI in the reference sheet
    if re.search(r'\bhedgehog\b|\bhedgehog\d+\b', str(npi_name).lower()):
        hedge_row = ref_df_shuffled[ref_df_shuffled['NPI NAME'].str.lower().str.contains(r'\bhedgehog\b', regex=True, na=False)]
        if not hedge_row.empty:
            return hedge_row.iloc[0]['NPI NAME'], hedge_row.iloc[0]['MODEL NUMBER']
            
    # Special-case: if 'hunter' appears anywhere in the extracted NPI name, reference the first 'Hunter' NPI in the reference sheet
    # Use pattern that matches "hunter" followed by optional number (like Hunter4, Hunter2, etc.)
    if re.search(r'\bhunter\b|\bhunter\d+\b', str(npi_name).lower()):
        hunter_row = ref_df_shuffled[ref_df_shuffled['NPI NAME'].str.lower().str.contains(r'\bhunter\b', regex=True, na=False)]
        if not hunter_row.empty:
            return hunter_row.iloc[0]['NPI NAME'], hunter_row.iloc[0]['MODEL NUMBER']
            
    # Special-case: if 'telluride' appears anywhere in the extracted NPI name, reference the first 'Telluride' NPI in the reference sheet
    if re.search(r'\btelluride\b|\btelluride\d+\b', str(npi_name).lower()):
        telluride_row = ref_df_shuffled[ref_df_shuffled['NPI NAME'].str.lower().str.contains(r'\btelluride\b', regex=True, na=False)]
        if not telluride_row.empty:
            return telluride_row.iloc[0]['NPI NAME'], telluride_row.iloc[0]['MODEL NUMBER']
            
    # Special-case: if 'pyrite' appears anywhere in the extracted NPI name, reference the first 'Pyrite' NPI in the reference sheet
    if re.search(r'\bpyrite\b|\bpyrite\d+\b', str(npi_name).lower()):
        pyrite_row = ref_df_shuffled[ref_df_shuffled['NPI NAME'].str.lower().str.contains(r'\bpyrite\b', regex=True, na=False)]
        if not pyrite_row.empty:
            return pyrite_row.iloc[0]['NPI NAME'], pyrite_row.iloc[0]['MODEL NUMBER']
            
    # Special-case: if 'hawk' appears as a standalone word in the extracted NPI name, reference the first 'Hawk' NPI in the reference sheet
    # But ONLY if the NPI name doesn't contain 'blackhawk' which is handled above
    if (re.search(r'\bhawk\b|\bhawk\d+\b', str(npi_name).lower()) and 
        not re.search(r'\bblackhawk\b|\bblackhawk\d+\b', str(npi_name).lower())):
        hawk_row = ref_df_shuffled[ref_df_shuffled['NPI NAME'].str.lower().str.contains(r'\bhawk\b', regex=True, na=False)]
        if not hawk_row.empty:
            return hawk_row.iloc[0]['NPI NAME'], hawk_row.iloc[0]['MODEL NUMBER']
    
    npi_name = str(npi_name).strip().lower()
    
    # Split and clean input NPI name
    input_parts = set()
    for part in npi_name.replace('/', ',').split(','):
        cleaned = ' '.join(p.strip() for p in part.split())
        if cleaned:
            input_parts.add(cleaned)
            input_parts.update(cleaned.split())
            
    # For short NPIs (<=2 tokens), require the main word (not a number) to appear in the reference NPI name
    input_tokens = [t for t in input_parts if t and not t.isdigit()]
    main_token = input_tokens[0] if input_tokens else None

    def _simplify_npi_part(s: str) -> str:
        s = str(s).lower()
        noise = ['pcb', 'board', 'pod', 'assembly', 'fixture', 'ap1', 'ap2', 'module', 'pv', 'deskew', 'mmcx']
        s = re.sub(r"[^0-9a-z ]", " ", s)
        tokens = [t for t in s.split() if t and t not in noise]
        return " ".join(tokens)

    simplified_input_parts = {_simplify_npi_part(p) for p in input_parts if p}
    
    variations = {
        'cps': ['cps', 'cps product', 'cps transfer', 'cps special handling'],
        'product': ['product', 'prod'],
        'transfer': ['transfer', 'move'],
        'special': ['special', 'spec'],
        'handling': ['handling', 'hand']
    }
    
    expanded_parts = set(input_parts)
    for part in input_parts:
        for var_list in variations.values():
            if part in var_list:
                expanded_parts.update(var_list)
    
    candidates = []
    for _, row in ref_df_shuffled.iterrows():
        ref_npi = str(row["NPI NAME"]).strip().lower()
        
        ref_parts = set()
        for part in ref_npi.replace('/', ',').split(','):
            cleaned = ' '.join(p.strip() for p in part.split())
            if cleaned:
                ref_parts.add(cleaned)
                ref_parts.update(cleaned.split())

        # Simplified reference parts for fuzzy/substring matching
        simplified_ref_parts = {_simplify_npi_part(p) for p in ref_parts if p}
        
        # Add variations to reference parts
        expanded_ref_parts = set(ref_parts)
        for part in ref_parts:
            for var_list in variations.values():
                if part in var_list:
                    expanded_ref_parts.update(var_list)
        
        direct_matches = expanded_parts & expanded_ref_parts
        # For short NPIs, require main token to be present in the reference NPI name
        if main_token and len(input_parts) <= 2:
            ref_npi_tokens = set(ref_npi.split())
            if not any(main_token.lower() == t.lower() for t in ref_npi_tokens):
                continue
        if not direct_matches:
            simp_direct = any(
                (sip and any(sip == srp or sip in srp or srp in sip for srp in simplified_ref_parts))
                for sip in simplified_input_parts
            )
            if simp_direct:
                direct_matches = {d for d in expanded_parts if _simplify_npi_part(d) in simplified_ref_parts}
            else:
                fuzzy_hits = 0
                for sip in simplified_input_parts:
                    if not sip or len(sip) < 3:
                        continue
                    best = 0.0
                    for srp in simplified_ref_parts:
                        if not srp:
                            continue
                        score = SequenceMatcher(None, sip, srp).ratio()
                        if score > best:
                            best = score
                    if best >= 0.82:
                        fuzzy_hits += 1
                if fuzzy_hits > 0:
                    direct_matches = {next(iter(expanded_parts))}
                else:
                    continue

        exact_matches = len(direct_matches)
        partial_matches = sum(1 for p in expanded_parts for r in expanded_ref_parts if p in r or r in p)

        key_terms = {'cps', 'product', 'transfer', 'special', 'handling'}
        key_matches = sum(1 for term in key_terms if any(term in p for p in direct_matches))

        if len(input_parts) <= 2 and exact_matches >= 1:
            confidence = 0.85
            candidates.append((confidence, len(direct_matches), len(ref_npi), row))
            continue

        confidence = (
            (exact_matches / len(expanded_parts)) * 0.45 +
            (partial_matches / (len(expanded_parts) + len(expanded_ref_parts))) * 0.25 +
            (key_matches / max(1, len(key_terms))) * 0.30
        )

        if confidence >= 0.3:
            candidates.append((confidence, len(direct_matches), len(ref_npi), row))
    
    if candidates:
        candidates.sort(key=lambda x: (x[0], x[1], -x[2]), reverse=True)
        best_match = candidates[0]
        if (best_match[0] >= 0.5 or 
            (len(candidates) > 1 and best_match[0] - candidates[1][0] >= 0.2)):
            return best_match[3]["NPI NAME"], best_match[3]["MODEL NUMBER"]
    
    return None, None

def _extract_model_family(model):
    """Extract the model family (e.g., 'N52' from 'N5247BC')."""
    # Match letter prefix and 2-3 digits
    match = re.match(r'^([A-Z]+\d{2,3})', model.upper())
    return match.group(1) if match else None

def _models_match(model, pattern):
    """Check if a model matches a pattern, including wildcards."""
    model = model.strip().upper()
    pattern = pattern.strip().upper()
    
    # Exact match
    if model == pattern:
        return True, 1.0
        
    # Handle wildcard patterns (e.g., N52xxx)
    if 'X' in pattern:
        base_pattern = pattern.rstrip('X')
        if model.startswith(base_pattern):
            return True, 0.9
    
    # Handle family matching
    model_family = _extract_model_family(model)
    pattern_family = _extract_model_family(pattern)
    
    if model_family and pattern_family:
        if model_family == pattern_family:
            return True, 0.8
    
    # No match
    return False, 0.0

def _extract_model_parts(model_str):
    """Split a model string into individual cleaned parts."""
    parts = []
    # Split by comma and slash
    for part in str(model_str).replace('/', ',').split(','):
        clean = re.sub(r'[^A-Z0-9X-]', '', part.strip().upper())
        if clean:
            parts.append(clean)
    return parts

def _get_model_family(model):
    """Extract model family (e.g., 'N52' from 'N5247BC')."""
    # Try to match alphanumeric prefix pattern (letters followed by digits)
    match = re.match(r'^([A-Z]+\d{2,3})', model)
    if match:
        return match.group(1)
    # Special case for models that start with digits (e.g., 12345A)
    digit_match = re.match(r'^(\d{2,4})', model)
    return digit_match.group(1) if digit_match else None

def _extended_model_match(model, pattern):
    """More flexible model matching that handles partial prefix matches."""
    model = model.strip().upper()
    pattern = pattern.strip().upper()
    
    # First check standard matching
    matched, score = _models_match(model, pattern)
    if matched:
        return True, score
        
    # Check for shared prefixes (more flexible matching)
    # Get the alphanumeric parts of each model
    model_prefix = re.match(r'^([A-Z]*\d+)', model)
    pattern_prefix = re.match(r'^([A-Z]*\d+)', pattern)
    
    if model_prefix and pattern_prefix:
        m_prefix = model_prefix.group(1)
        p_prefix = pattern_prefix.group(1)
        
        # Check if they share a significant prefix (at least 4 chars)
        shared_len = 0
        for i in range(min(len(m_prefix), len(p_prefix))):
            if m_prefix[i] == p_prefix[i]:
                shared_len += 1
            else:
                break
                
        if shared_len >= 4:  # Significant prefix match
            return True, 0.7 * (shared_len / max(len(m_prefix), len(p_prefix)))
    
    return False, 0.0

def match_by_model(model_number):
    """Match by model number with split comparison approach."""
    if _is_invalid_extracted(model_number):
        return None, None

    # Split input into individual models
    input_models = _extract_model_parts(model_number)
    if not input_models:
        return None, None

    candidates = []
    for _, row in ref_df_shuffled.iterrows():
        ref_models = _extract_model_parts(row["MODEL NUMBER"])
        if not ref_models:
            continue

        matches = []
        ref_families = {_get_model_family(m) for m in ref_models if _get_model_family(m)}

        for input_model in input_models:
            best_match_score = 0
            match_found = False

            # Exact / pattern matches
            for ref_model in ref_models:
                # Use extended model matching for more flexible comparison
                matched, score = _extended_model_match(input_model, ref_model)
                if matched:
                    match_found = True
                    best_match_score = max(best_match_score, score)

            # Family fallback
            if not match_found:
                input_family = _get_model_family(input_model)
                if input_family and input_family in ref_families:
                    match_found = True
                    best_match_score = 0.7

            if match_found:
                matches.append(best_match_score)

        if matches:
            match_ratio = len(matches) / len(input_models)
            avg_match_score = sum(matches) / len(matches)
            # Increase weight of match score for better precision
            confidence = (match_ratio * 0.5) + (avg_match_score * 0.5)
            if confidence >= 0.3:
                candidates.append((confidence, len(matches), row))

    if candidates:
        candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
        best = candidates[0]
        if (best[0] >= 0.5 or
            (len(candidates) > 1 and best[0] - candidates[1][0] >= 0.2)):
            return best[2]["NPI NAME"], best[2]["MODEL NUMBER"]

    return None, None

# Validate required columns for ISO document check
required_columns = ["CATEGORY OF CHANGE", "CHANGE COORDINATOR"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    print("⚠️ Warning: Some columns required for ISO document check are missing:", missing_columns)
    has_iso_columns = False
else:
    has_iso_columns = True

# Init columns
df["Referenced NPI Name"] = None
df["Referenced Model Number"] = None
df["Product Line"] = df.get("Product Line", None)  # Only initialize if doesn't exist
df["Business Group"] = df.get("Business Group", None)  # Only initialize if doesn't exist

# Special case: ISO Document and Attribute Only handling
if has_iso_columns:
    # ISO Document case
    iso_mask = (
        (df["CATEGORY OF CHANGE"].str.strip().str.upper() == "ISO DOCUMENT") & 
        (df["CHANGE COORDINATOR"].str.strip() == "sin-hui.teh@non.keysight.com")
    )
    df.loc[iso_mask, "Referenced NPI Name"] = "TIS Back End"
    df.loc[iso_mask, "Referenced Model Number"] = "CCD"
    df.loc[iso_mask, "Product Line"] = "TIS"
    df.loc[iso_mask, "Business Group"] = "N/A"
    
    # Attribute Only case
    attribute_mask = (
        (df["CATEGORY OF CHANGE"].str.strip().str.upper() == "ATTRIBUTE ONLY") & 
        (df["CHANGE COORDINATOR"].str.strip() == "sin-hui.teh@non.keysight.com")
    )
    df.loc[attribute_mask, "Referenced NPI Name"] = "TIS Back End"
    df.loc[attribute_mask, "Referenced Model Number"] = "DG"
    df.loc[attribute_mask, "Product Line"] = "TIS"
    df.loc[attribute_mask, "Business Group"] = "N/A"

# Pass 1: Match each row independently to prevent cascading effects
for idx, row in df.iterrows():
    matched = False
    
    # Try NPI name first (prioritize NPI name matching)
    if not _is_invalid_extracted(row["NPI Name"]):
        npi_ref, model_ref = match_by_npi(row["NPI Name"])
        if npi_ref and not _is_invalid_extracted(model_ref):
            df.at[idx, "Referenced NPI Name"] = npi_ref
            df.at[idx, "Referenced Model Number"] = model_ref
            matched = True
    
    # If NPI didn't match, try model number
    if not matched and not _is_invalid_extracted(row["Model Number"]):
        npi_ref, model_ref = match_by_model(row["Model Number"])
        if model_ref:
            # Verify the match - both model and NPI should be non-empty
            if not _is_invalid_extracted(npi_ref):
                df.at[idx, "Referenced NPI Name"] = npi_ref
                df.at[idx, "Referenced Model Number"] = model_ref

# ===== Pass 2: Fill missing from counterpart =====
for idx, row in df.iterrows():
    if pd.isna(row["Referenced NPI Name"]) and pd.notna(row["Referenced Model Number"]):
        match_row = ref_df_shuffled[ref_df_shuffled["MODEL NUMBER"].apply(lambda cell: _model_in_cell(cell, row["Referenced Model Number"]))]
        if not match_row.empty:
            df.at[idx, "Referenced NPI Name"] = match_row.iloc[0]["NPI NAME"]

    if pd.isna(row["Referenced Model Number"]) and pd.notna(row["Referenced NPI Name"]):
        match_row = ref_df_shuffled[ref_df_shuffled["NPI NAME"].str.strip().str.lower() ==
                           str(row["Referenced NPI Name"]).strip().lower()]
        if not match_row.empty:
            df.at[idx, "Referenced Model Number"] = match_row.iloc[0]["MODEL NUMBER"]

# ===== STEP 6 (pre): Refine 'WN' Product Line using ReferenceModelNumber =====
# If extracted Product Line is the generic 'WN', use referenced identifiers to get the specific WN variant.
if "Product Line" in df.columns:
    # Column name helpers
    ref_colname_map = {c.lower(): c for c in ref_df.columns}
    product_line_col = (
        ref_colname_map.get("product line")
        or ref_colname_map.get("product_line")
        or ref_colname_map.get("productline")
        or ref_colname_map.get("pl name")
        or ref_colname_map.get("pl_name")
        or ref_colname_map.get("pl")
    )
    if product_line_col is None:
        # Fallback: any column whose name contains both 'product' and 'line'
        for col in ref_df.columns:
            low = str(col).lower()
            if "product" in low and "line" in low:
                product_line_col = col
                break

    def _find_ref_row(row):
        # Prefer robust best-match lookups to handle minor formatting differences
        model_col = ref_colname_map.get("model number") or "MODEL NUMBER"
        npi_col = ref_colname_map.get("npi name") or "NPI NAME"
        # Try keys in order of reliability: Referenced Model, Referenced NPI, Extracted Model, Extracted NPI
        keys = [
            ("Referenced Model Number", model_col, True),
            ("Referenced NPI Name", npi_col, False),
            ("Model Number", model_col, True),
            ("NPI Name", npi_col, False),
        ]
        for df_key, ref_key, prefer_token in keys:
            val = row.get(df_key)
            if _is_invalid_extracted(val):
                continue
            # If matching by model, narrow rows to those whose 'MODEL NUMBER' cell contains the model token
            if prefer_token and ref_key == model_col:
                narrowed_df = ref_df[ref_df[ref_key].apply(lambda cell: _model_in_cell(cell, val))]
                if not narrowed_df.empty:
                    bm_row = _best_match_row(val, narrowed_df, ref_key, fuzzy_threshold=0.82, prefer_token=prefer_token)
                    if bm_row is not None:
                        return bm_row
            bm_row = _best_match_row(val, ref_df, ref_key, fuzzy_threshold=0.82, prefer_token=prefer_token)
            if bm_row is not None:
                return bm_row
        return None

    def _extract_specific_wn_from_ref_row(ref_row):
        # Hardcoded special-case: model list corresponds to 'WN Kobe'
        try:
            model_col = ref_colname_map.get("model number") or "MODEL NUMBER"
            model_cell = str(ref_row.get(model_col, ""))
            cell_tokens = {_normalize(t) for t in _tokens(model_cell)}
            target_models = [
                "E4980A", "E4980B", "E4981B", "E4980", "E4981BU",
                "E5054", "E5055A", "E5056A", "E5057A", "E5058A", "E5052B", "S963105B",
            ]
            target_tokens = {_normalize(t) for t in target_models}
            if cell_tokens & target_tokens:
                return "WN Kobe"
        except Exception:
            pass

        # First prefer explicit product line column value
        if product_line_col is not None:
            cand = str(ref_row.get(product_line_col, "")).strip()
            if cand and cand.upper().startswith("WN") and len(cand) > 2:
                return cand
        # Otherwise scan the row for a 'WN <type>' token
        for col in ref_row.index:
            s = str(ref_row[col]).strip()
            if not s:
                continue
            m = re.search(r"\bWN\b\s*[:\-_/]?\s*([^,;\n\r]+)", s, flags=re.IGNORECASE)
            if m:
                wn_type = m.group(1).strip()
                if wn_type and wn_type.upper() != "WN":
                    return f"WN {wn_type}"
            if s.upper().startswith("WN") and len(s) > 2:
                return s
        # Fallback: derive from platform/family/program-like fields by prefixing 'WN '
        candidate_keys = [
            "wn type", "wn_type", "type", "subtype", "pl type", "pl_type", "pl name", "pl_name",
            "platform", "family", "program", "category", "segment"
        ]
        # Build a map of lower-case column names to actual names
        lower_to_col = {str(c).lower(): c for c in ref_row.index}
        for key in candidate_keys:
            col = lower_to_col.get(key)
            if col is None:
                continue
            val = str(ref_row[col]).strip()
            if not val:
                continue
            u = val.upper()
            if u in {"WN", "N/A", "NA", "NONE", "NULL", "-"}:
                continue
            # Avoid double 'WN '
            if u.startswith("WN"):
                if len(val) > 2:
                    return val
                continue
            return f"WN {val}"
        return None

    def _needs_wn_refinement(pl_value) -> bool:
        if pd.isna(pl_value):
            return False
        s = str(pl_value)
        up = re.sub(r"\s+", " ", s.strip()).upper()
        # Already a known specific WN type? skip
        known_specifics = {"WN KOBE", "WN PMPS", "WN SOCO", "WN TAO"}
        if any(up.startswith(spec) for spec in known_specifics):
            return False
        # Is there a 'WN' token present? covers 'WN', '-WN', 'WN-', '... WN ...'
        return re.search(r"(^|[^A-Z0-9])WN([^A-Z0-9]|$)", up) is not None

    for idx, row in df.iterrows():
        current_pl = row.get("Product Line")
        if _needs_wn_refinement(current_pl):
            # Look up the specific WN variant through reference data

            ref_row = _find_ref_row(row)
            if ref_row is None:
                continue
            specific_pl = _extract_specific_wn_from_ref_row(ref_row)
            if specific_pl and str(specific_pl).strip() and str(specific_pl).strip().upper().startswith("WN") and len(str(specific_pl).strip()) > 2:
                df.at[idx, "Product Line"] = str(specific_pl).strip()

# ===== STEP 6: Map Product Line → Business Group =====
# Transform Product Line "08" to "1A"
df.loc[df["Product Line"] == "08", "Product Line"] = "1A"

# Create mapping from Product Line to Business Group
pl_to_bg = dict(zip(mapping_df["PRODUCT LINE"], mapping_df["BUSINESS GROUP"]))
df["Business Group"] = df["Product Line"].map(pl_to_bg)

def partial_match(pl_value):
    if pd.isna(pl_value):
        return None
    pl_value = str(pl_value).lower().strip()
    
    # Transform "08" to "1A" in partial matching
    if pl_value == "08":
        pl_value = "1a"
        
    for _, row in mapping_df.iterrows():
        if str(row["PRODUCT LINE"]).lower().strip() in pl_value:
            return row["BUSINESS GROUP"]
    return None

df.loc[df["Business Group"].isna(), "Business Group"] = df.loc[df["Business Group"].isna(), "Product Line"].apply(partial_match)

# ===== STEP 6.2: Final backfill of Referenced fields by Product Line =====
# Only if both Referenced fields are still empty, try to fill them by matching Product Line to the reference sheet.
ref_colname_map = {c.lower().strip(): c for c in ref_df.columns}
# Identify the product line column in the reference sheet robustly
ref_product_line_col = (
    ref_colname_map.get("product line")
    or ref_colname_map.get("product_line")
    or ref_colname_map.get("productline")
)
if ref_product_line_col is None:
    # Try fuzzy detection: any column whose name contains both 'product' and 'line'
    for col in ref_df.columns:
        low = col.lower()
        if "product" in low and "line" in low:
            ref_product_line_col = col
            break
if ref_product_line_col is not None:
    model_col = ref_colname_map.get("model number") or "MODEL NUMBER"
    npi_col = ref_colname_map.get("npi name") or "NPI NAME"

    def _pl_norm(val: str) -> str:
        return _normalize(val)

    for idx, row in df.iterrows():
        need_ref_npi = _is_invalid_extracted(row.get("Referenced NPI Name"))
        need_ref_model = _is_invalid_extracted(row.get("Referenced Model Number"))
        if not (need_ref_npi and need_ref_model):
            continue
        pl_val = row.get("Product Line")
        if _is_invalid_extracted(pl_val):
            continue

        pl_norm = _pl_norm(str(pl_val))

        # Build candidates where normalized PLs are equal or contain one another;
        # If the PL is generic 'wn', allow any reference PL that starts with 'wn'
        ref_pl_series = ref_df[ref_product_line_col].astype(str)
        def _candidate_mask(x: str) -> bool:
            x_norm = _pl_norm(x)
            if not x_norm:
                return False
            if pl_norm == x_norm or pl_norm in x_norm or x_norm in pl_norm:
                return True
            if pl_norm == "wn":
                return x_norm.startswith("wn")
            return False

        ref_candidates = ref_df[ref_pl_series.apply(_candidate_mask)]
        if ref_candidates.empty:
            continue

        # If multiple candidates, prefer the one that best matches the extracted Model Number or NPI Name
        candidate_row = None
        extracted_model = row.get("Model Number")
        extracted_npi = row.get("NPI Name")
        if not _is_invalid_extracted(extracted_model):
            bm = _best_match_row(extracted_model, ref_candidates, model_col, fuzzy_threshold=0.82, prefer_token=True)
            if bm is not None:
                candidate_row = bm
        if candidate_row is None and not _is_invalid_extracted(extracted_npi):
            bm = _best_match_row(extracted_npi, ref_candidates, npi_col, fuzzy_threshold=0.82, prefer_token=False)
            if bm is not None:
                candidate_row = bm
        if candidate_row is None:
            candidate_row = ref_candidates.iloc[0]

        cand_model = str(candidate_row[model_col]).strip() if model_col in candidate_row else None
        cand_npi = str(candidate_row[npi_col]).strip() if npi_col in candidate_row else None
        if cand_model:
            df.at[idx, "Referenced Model Number"] = cand_model
        if cand_npi:
            df.at[idx, "Referenced NPI Name"] = cand_npi

# ===== STEP 6.3: Fill missing Product Line by matching Referenced NPI/Model against ReferenceModelNumber =====
# Runs at the end so it won't affect earlier logic.
ref_colname_map = {c.lower().strip(): c for c in ref_df.columns}
ref_product_line_col = (
    ref_colname_map.get("product line")
    or ref_colname_map.get("product_line")
    or ref_colname_map.get("productline")
)
if ref_product_line_col is None:
    for col in ref_df.columns:
        low = str(col).lower()
        if "product" in low and "line" in low:
            ref_product_line_col = col
            break

model_col = ref_colname_map.get("model number") or "MODEL NUMBER"
npi_col = ref_colname_map.get("npi name") or "NPI NAME"

def _get_pl_from_ref_row(r):
    if r is None:
        return None
    if ref_product_line_col is not None and ref_product_line_col in r.index:
        val = str(r[ref_product_line_col]).strip()
        return val if val else None
    return None

for idx, row in df[df["Product Line"].isna()].iterrows():
    candidate = None
    ref_model = row.get("Referenced Model Number")
    ref_npi = row.get("Referenced NPI Name")

    # Try model-based narrowing first
    if not pd.isna(ref_model):
        narrowed = ref_df[ref_df[model_col].apply(lambda cell: _model_in_cell(cell, ref_model))]
        if not narrowed.empty:
            # If NPI also present, refine within narrowed set
            if not pd.isna(ref_npi):
                bm = _best_match_row(ref_npi, narrowed, npi_col, fuzzy_threshold=0.82, prefer_token=False)
                if bm is not None:
                    candidate = bm
            if candidate is None:
                # Otherwise pick best by model within narrowed
                bm = _best_match_row(ref_model, narrowed, model_col, fuzzy_threshold=0.82, prefer_token=True)
                if bm is not None:
                    candidate = bm
            if candidate is None:
                candidate = narrowed.iloc[0]
        else:
            # No narrowed; try best-match on full table
            bm = _best_match_row(ref_model, ref_df, model_col, fuzzy_threshold=0.82, prefer_token=True)
            if bm is not None:
                candidate = bm

    # Fallback to NPI-only
    if candidate is None and not pd.isna(ref_npi):
        bm = _best_match_row(ref_npi, ref_df, npi_col, fuzzy_threshold=0.82, prefer_token=False)
        if bm is not None:
            candidate = bm

    pl_val = _get_pl_from_ref_row(candidate)
    if pl_val:
        df.at[idx, "Product Line"] = pl_val

# After adding PLs, try to fill Business Group for rows still missing
needs_bg = df["Business Group"].isna()
if needs_bg.any():
    # Transform any "08" to "1A" again in case new ones were added
    df.loc[df["Product Line"] == "08", "Product Line"] = "1A"
    
    # Try mapping again
    df.loc[needs_bg, "Business Group"] = df.loc[needs_bg, "Product Line"].map(pl_to_bg)
    still_missing = df["Business Group"].isna()
    if still_missing.any():
        df.loc[still_missing, "Business Group"] = df.loc[still_missing, "Product Line"].apply(partial_match)

# ===== STEP 7: Keep Final Columns =====
df_final = df[[
    desc_col,
    "NPI Name",
    "Model Number",
    "Referenced NPI Name",
    "Referenced Model Number",
    "Product Line",
    "Business Group"
]]

# ===== STEP 8: Save Output =====
df_final.to_excel(output_file, index=False)
print(f"✅ Extraction complete! File saved as {output_file}")
