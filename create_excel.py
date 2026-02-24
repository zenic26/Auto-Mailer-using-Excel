import pandas as pd

# Create sample data for the Excel file
data = {
    "Company": [
        "ABC Ltd",
        "XYZ Inc", 
        "Tech Co",
        "Global Solutions",
        "Innovate Tech"
    ],
    "Email": [
        "person1@gmail.com",
        "person2@gmail.com",
        "person3@gmail.com",
        "person4@gmail.com",
        "person5@gmail.com"
    ]
}

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel file
df.to_excel("emails.xlsx", index=False)

print("emails.xlsx created successfully!")
print("\nFile contents:")
print(df)
