import csv
import json
import urllib.request
import urllib.error
from collections import defaultdict

def make_request(method, url, headers, data=None):
    if data:
        data = json.dumps(data).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req) as response:
            return json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        print(f"Error: {e.code} {e.reason} {e.read().decode()}")
        raise

def get_sheet(token, sheet_id):
    url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    return make_request("GET", url, headers)

def batch_delete_rows(token, sheet_id, row_ids, batch_size=200):
    base_url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows"
    headers = {"Authorization": f"Bearer {token}"}
    for i in range(0, len(row_ids), batch_size):
        batch = row_ids[i:i+batch_size]
        ids_str = ','.join(map(str, batch))
        delete_url = f"{base_url}?ids={ids_str}"
        make_request("DELETE", delete_url, headers)

def batch_add_rows(token, sheet_id, new_rows, batch_size=400):
    base_url = f"https://api.smartsheet.com/2.0/sheets/{sheet_id}/rows"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    for i in range(0, len(new_rows), batch_size):
        batch = new_rows[i:i+batch_size]
        make_request("POST", base_url, headers, batch)

if __name__ == "__main__":
    token = "wQ24CHJ3HOFTtpUKku9i07esDdI8dD5R22SrY" # Get your API key using instructions at https://help.smartsheet.com/articles/2482389-generate-API-key
    sheet_id = 3699213415174020 #Find your Sheet ID via File / Properties menu
    csv1_file = "offices.csv"
    csv2_file = "users.csv"

    # Read CSV1
    entities = {}
    with open(csv1_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            loc = row.get("Location Code")
            if loc:
                entities[loc] = row

    # Read CSV2
    users = defaultdict(lambda: defaultdict(list))
    with open(csv2_file, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            loc = row.get("Location Code")
            email = row.get("Email Address")
            role = row.get("Business Role")
            if loc and email and role:
                users[loc][role].append(email)

    # Get sheet
    sheet = get_sheet(token, sheet_id)
    title_to_id = {col["title"]: col["id"] for col in sheet["columns"]}

    # Delete all existing rows in batches
    row_ids = [r["id"] for r in sheet.get("rows", [])]
    if row_ids:
        batch_delete_rows(token, sheet_id, row_ids)

    # Build new rows
    new_rows = []
    multi_cols = ["Brands"]
    for loc, data in entities.items():
        cells = []
        for col, val in data.items():
            col_id = title_to_id.get(col)
            if not col_id:
                print(f"No column for {col}")
                continue
            if col in multi_cols:
                values = [v.strip() for v in val.split(',') if v.strip()]
                if values:
                    ov = {"objectType": "MULTI_PICKLIST", "values": values}
                    cells.append({"columnId": col_id, "objectValue": ov})
            else:
                cells.append({"columnId": col_id, "value": val})

        # Add contacts from CSV2
        for role, emails in users.get(loc, {}).items():
            col_id = title_to_id.get(role)
            if col_id:
                if emails:
                    values = [{"objectType": "CONTACT", "email": e} for e in emails]
                    ov = {"objectType": "MULTI_CONTACT", "values": values}
                    cells.append({"columnId": col_id, "objectValue": ov})
            else:
                print(f"No column for role: {role}")

        if cells:
            new_rows.append({"toBottom": True, "cells": cells})

    # Add new rows in batches
    if new_rows:
        batch_add_rows(token, sheet_id, new_rows)
