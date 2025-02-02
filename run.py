import pandas as pd

from torob.main import Torob


def main():
    # List of search queries
    search_queries = [
        "لپ تاپ",
        "موبایل",
    ]

    # Create an Excel writer object
    file_path = "all_products.xlsx"
    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        with Torob(teardown=True) as bot:
            for query in search_queries:
                bot.land_first_page()
                bot.search_box(query)
                bot.sort_items(text="محبوب")
                bot.find_commodities()
                df = bot.write_to_file()
                # Write the DataFrame to a sheet named after the query
                df.to_excel(writer, sheet_name=query, index=False)
                print(f"Data for '{query}' written to sheet '{query}' in {file_path}")


if __name__ == "__main__":
    main()
