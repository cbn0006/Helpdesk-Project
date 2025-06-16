import pandas as pd


def main():
    rawCSV = pd.read_csv('Current.csv')

    rawCSV.head()

    newFile = input("Please enter new file's name: ")


if __name__ == "__main__":
    main()
    