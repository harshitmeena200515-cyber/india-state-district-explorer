import json

with open("states-and-districts.json", "r", encoding="utf-8") as file:
    data = json.load(file)

state_name = input("Enter State Name: ")

found = False

for item in data["states"]:
    if item["state"].lower() == state_name.lower():
        print(f"\nDistricts of {item['state']}:\n")

        for i, district in enumerate(item["districts"], start=1):
            print(f"{i}: {district}")

        found = True
        break

if not found:
    print("State not found")
