import pandas as pd
from rapidfuzz import process, fuzz

# Load data
input_file = "Name matching.csv"
output_file = "Name matching with auto number.xlsx"

df = pd.read_csv(input_file)

# Separate baseline and endline
baseline = df[df["Timepoint"].str.lower() == "baseline"].reset_index(drop=True)
endline = df[df["Timepoint"].str.lower() == "endline"].reset_index(drop=True)

# Prepare for matching
endline_available = set(endline.index)
baseline['auto number'] = None
endline['auto number'] = None

auto_id = 1

for i, base_row in baseline.iterrows():
    base_name = base_row['Full Name']
    # Find best fuzzy match among unmatched endline names
    if endline_available:
        matches = process.extract(
            base_name,
            endline.loc[list(endline_available), "Full Name"],
            scorer=fuzz.token_sort_ratio,
            limit=1
        )
        if matches and matches[0][1] >= 85:  # You can adjust threshold here
            match_name, score, match_idx = matches[0]
            baseline.at[i, 'auto number'] = auto_id
            endline.at[match_idx, 'auto number'] = auto_id
            endline_available.remove(match_idx)
            auto_id += 1

# Assign unique IDs to unmatched baseline rows
for i, row in baseline[baseline['auto number'].isna()].iterrows():
    baseline.at[i, 'auto number'] = auto_id
    auto_id += 1

# Assign unique IDs to unmatched endline rows
for i, row in endline[endline['auto number'].isna()].iterrows():
    endline.at[i, 'auto number'] = auto_id
    auto_id += 1

# Concatenate and save
result = pd.concat([baseline, endline], ignore_index=True)
result.to_excel(output_file, index=False)
print(f"Done! Output saved to {output_file}")
