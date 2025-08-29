import requests
import random
import os
import pandas as pd

def create_name_list():
    url = "https://api.mir6.com/api/sjname"
    out = {}
    name = {}
    sex = {}
    row = []
    col = []
    
    numbers = input("Enter the number of names to generate (default is 100): ")
    if not numbers.isdigit():
        numbers = 100
    

    for i in range(int(numbers)):
        response = requests.get(url)
        if response.status_code == 200:
            line = response.text + "," + random.choice(["男", "女"]) + "," + str(random.randint(155, 175)) + "," + str(i+1) + "," + str(random.randint(60, 100))
            row = line.split(",")
            col.append(row)
            print(line)

            
        

        else:
            print("Failed to retrieve name from API, using default name.")
    df = pd.DataFrame(col, columns=["name", "gender", "height(cm)", "number", "marks"])
    df.to_csv("name_list.csv", index=False, encoding="utf-8")
    df.to_excel("name_list.xlsx", index=False, sheet_name="Sheet1")
    df.to_csv("name_list.txt", index=False, encoding="utf-8", sep=" ")

    

    
if __name__ == "__main__":
    create_name_list()