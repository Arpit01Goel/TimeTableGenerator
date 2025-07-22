import pandas as pd
from bs4 import BeautifulSoup
from collections import defaultdict

# STEP 1: Read Excel
filename = "TimeTable.xlsx"
sheet_name = "Time Table"

# Skip first 3 rows (metadata), read actual timetable
df_raw = pd.read_excel(filename, sheet_name=sheet_name, header=None, skiprows=3)

# STEP 2: Input list of required course codes
input_str = input("Enter course codes to include (comma-separated): ")
selected_courses = [code.strip().upper() for code in input_str.split(",")]

# STEP 3: Prepare headers
time_slots = df_raw.iloc[0, 1:].tolist()
df = df_raw.iloc[1:].copy()
df.columns = ['Day'] + time_slots

# STEP 4: Group rows by day, collapse entries by time slot
grouped_data = defaultdict(lambda: {slot: [] for slot in time_slots})

current_day = None
for i, row in df.iterrows():
    day = row['Day']
    if pd.notna(day):
        current_day = day.strip()
    if current_day is None:
        continue
    for slot in time_slots:
        cell = row[slot]
        if pd.isna(cell):
            continue
        entries = str(cell).split('\n') if '\n' in str(cell) else [cell]
        for entry in entries:
            for course in selected_courses:
                if course in entry:
                    parts = entry.split('(')
                    code = parts[0].strip()
                    room = '(' + parts[-1] if len(parts) > 1 else ''
                    formatted = f"{code}<br><small>{room}</small>"
                    grouped_data[current_day][slot].append(formatted)

# STEP 5: Generate new clean HTML
html = f"""
<html>
<head>
    <style>
        body {{
            font-family: Arial, sans-serif;
            padding: 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        th, td {{
            border: 1px solid #aaa;
            padding: 8px;
            text-align: center;
            vertical-align: middle;
            min-width: 90px;
        }}
        th {{
            background-color: #f0f0f0;
        }}
        td small {{
            color: #555;
        }}
    </style>
</head>
<body>
    <h2>Your Customized Timetable</h2>
    <table>
        <tr>
            <th>Day</th>
"""

for slot in time_slots:
    html += f"<th>{slot}</th>"
html += "</tr>\n"

for day, slots in grouped_data.items():
    html += f"<tr><td><b>{day}</b></td>"
    for slot in time_slots:
        cell_content = "<hr>".join(slots[slot]) if slots[slot] else ""
        html += f"<td>{cell_content}</td>"
    html += "</tr>\n"

html += """
    </table>
</body>
</html>
"""

# Allow manual edits to timetable
while True:
    print("\nDo you want to:")
    print("1. Add a class manually")
    print("2. Remove a class manually")
    print("3. Finish and generate timetable")
    choice = input("Enter option (1/2/3): ").strip()

    if choice == "1":
        day = input("Day (e.g. Monday): ").strip()
        slot = input("Time slot (exactly as shown, e.g. 9.00 - 9.50): ").strip()
        code = input("Course code (e.g. CS303): ").strip()
        room = input("Room (e.g. M6): ").strip()
        formatted = f"{code}<br><small>({room})</small>"
        grouped_data[day][slot].append(formatted)
        print(f"✅ Added {code} at {slot} on {day}")

    elif choice == "2":
        day = input("Day (e.g. Monday): ").strip()
        slot = input("Time slot (e.g. 9.00 - 9.50): ").strip()
        code = input("Course code to remove (e.g. CS303): ").strip()
        new_entries = [entry for entry in grouped_data[day][slot] if code not in entry]
        grouped_data[day][slot] = new_entries
        print(f"🗑️ Removed {code} from {slot} on {day}")

    elif choice == "3":
        break
    else:
        print("❌ Invalid option. Please enter 1, 2, or 3.")


with open("my_timetable.html", "w", encoding="utf-8") as f:
    f.write(html)

print("✅ Timetable generated: my_timetable.html")
