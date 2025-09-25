import pandas as pd
from rapidfuzz.fuzz import token_set_ratio, partial_ratio
from rapidfuzz import process

def normalize(name):
    """Lowercase and remove extra spaces from a name."""
    return ' '.join(str(name).lower().strip().split())

# === CONFIGURATION ===
BASELINE_FILE = "Name matching 2.csv"  # Change to your input file
ENDLINE_COLUMN = "Endline"             # Change to your endline column name
BASELINE_COLUMN = "Baseline"           # Change to your baseline column name
THRESHOLD = 85                         # Adjust this threshold if needed

# === LOAD DATA ===
df = pd.read_csv(BASELINE_FILE)

# If you have a single file with both columns:
baseline_names = df[BASELINE_COLUMN].dropna().unique()
endline_names = df[ENDLINE_COLUMN].dropna().unique()

# === MATCHING LOGIC ===
results = []
for base in baseline_names:
    norm_base = normalize(base)
    best_match = None
    best_score = 0
    for end in endline_names:
        norm_end = normalize(end)
        score = max(
            token_set_ratio(norm_base, norm_end),
            partial_ratio(norm_base, norm_end)
        )
        if score > best_score:
            best_score = score
            best_match = end
    match_status = "MATCH" if best_score >= THRESHOLD else "NO MATCH"
    results.append({
        BASELINE_COLUMN: base,
        ENDLINE_COLUMN: best_match if best_score >= THRESHOLD else "",
        "Score": best_score,
        "Status": match_status
    })

# === OUTPUT RESULTS ===
output_df = pd.DataFrame(results)
output_file = BASELINE_FILE.replace(".csv", " with matches.xlsx")
output_df.to_excel(output_file, index=False)
print(f"Matching complete. Results saved to {output_file}")
