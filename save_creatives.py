import pandas as pd
import requests
import sys
import os


def verify_input():
    if len(sys.argv) < 2:
        files = os.listdir()
        for file in files:
            if file.endswith(".xlsx"):
                file = file
                print(f"using {file}")
                return file
    else:
        file = sys.argv[1]
        if not file.endswith(".xlsx"):
            sys.exit("Input file must be an excel file (.xlsx)")

    sys.exit("No excel file")


def make_entry(row):
    entry = {}
    entry["link"] = row[0]
    entry["headline"] = row[1]
    entry["campaign"] = row[2]
    entry["creatives"] = [row[3], row[4], row[5]]

    return entry


def save_creatives(entry):
    # download images to folder
    for image in entry["creatives"]:
        r = requests.get(image)
        # save the image to directory
        with open(image.split("/")[-1], "wb") as f:
            f.write(r.content)

    # make a text file named info.txt
    with open("info.txt", "w") as f:
        f.write(f"Headline: {entry['headline']}\n")
        f.write(f"Link: {entry['link']}\n")
        f.write(f"Campaigns: {entry['campaign']}")


def main():
    file = verify_input()
    df = pd.read_excel(file)

    registry = []
    for index, row in df.iterrows():
        entry = make_entry(row)
        registry.append(entry)

    counter = 0
    for entry in registry:
        counter += 1
        dir = entry["campaign"]
        # create folder for campaign
        try:
            os.mkdir(dir)
            os.chdir(dir)

            save_creatives(entry)
            # go back to parent directory
            os.chdir("..")
            print(f"{counter}/{len(registry)}")

            # clear the screen
            print("")
            print("\033c", end="")
        except FileExistsError:
            print(f"{dir} already exists")
            continue


if __name__ == "__main__":
    main()
