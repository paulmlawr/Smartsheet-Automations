import csv
from collections import defaultdict

# File paths
offices_file = 'offices.csv'  # CSV 1
users_file = 'users.csv'      # CSV 2
output_file = 'combined_offices.csv'

# Step 1: Read users and organize by Location Code and Role
location_role_email = defaultdict(dict)
roles_set = set()

with open(users_file, newline='', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        location_code = row['Location Code'].strip() #location_code is the key value in both csv files.
        role = row['Role'].strip()
        email = row['Email'].strip()
        location_role_email[location_code][role] = email
        roles_set.add(role)

# Step 2: Read office profiles
with open(offices_file, newline='', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    office_rows = list(reader)
    office_fieldnames = reader.fieldnames if reader.fieldnames else []

# Step 3: Prepare output fieldnames
sorted_roles = sorted(roles_set)  # Optional: consistent column order

# Reorder to put 'Location Code' first
other_office_fields = [f for f in office_fieldnames if f != 'Location Code']
combined_fieldnames = ['Location Code'] + other_office_fields + sorted_roles

# Step 4: Write output
with open(output_file, 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=combined_fieldnames)
    writer.writeheader()

    for office in office_rows:
        location_code = office['Location Code'].strip()
        role_emails = location_role_email.get(location_code, {})

        # Add role columns
        for role in sorted_roles:
            office[role] = role_emails.get(role, 'Not Found')

        writer.writerow(office)

print(f"Combined CSV written to '{output_file}' with 'Location Code' as the first column.")
