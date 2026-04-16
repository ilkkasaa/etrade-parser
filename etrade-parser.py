# Etrade Excel file parser.

# requires openpyxl and pandas:
# pip install openpyxl pandas

import argparse
import os
import pandas as pd
from datetime import datetime, timedelta

def parse_args() -> tuple[str, str]:
    help_text = """"
    You need to provide the Etrade Excel file and the valuuttakurssit csv file. 
    The Etrade Excel file can be downloaded from Etrade website:
    - At work -> My Account -> Gains & Losses.
    - Select the wanted tax year or custom time interval and click 'Apply'.
    - Click 'Download' and select 'Download Expanded'.

    Note: this list contains only shares sold in the selected year.

    Additionally you may need to manually check the currency exchange rates for ESPP purchases, if you have sold any ESPP shares during the year.
    The script will ask for the exchange rate for each ESPP purchase.

    Download the valuuttakurssit csv file from the Suomen Pankki website:
    https://www.suomenpankki.fi/fi/tilastot/taulukot-ja-kuviot/valuuttakurssit/taulukot/valuuttakurssit_taulukot_fi/valuuttakurssit_short_fi/
    'Kurssit': 'Yhdysvaltain dollari'. Make sure this is the only selection.
    'Alkupäivä': The first date of the period you want to download. Just make sure this is before the first date of the Etrade file.
    'Loppupäivä': The last date of the period you want to download. Just make sure this is after the last date of the Etrade file.
    'Uusimmat': 'ylimpänä'
    'Aggregointi': 'Keskiarvo'
    'Näytä tiedot': 'päivätasolla'

    Click save icon (export) and select 'CSV' as the file format.

    Note: this script does not give you any information about the dividends or potential gains/losses from any currency exchange.
    """
    parser = argparse.ArgumentParser(description='Parse Etrade Excel Files')
    parser.add_argument('-i', '--input', type=str, help='Etrade Excel File (Gains and Losses)', required=True)
    parser.add_argument('-v', '--valuuttakurssit', type=str, help='Valuuttakurssit csv file', required=True)

    args = parser.parse_args()

    # Check if files exist
    if not os.path.isfile(args.input):
        print("Input file does not exist.")
        return
    
    if not os.path.isfile(args.valuuttakurssit):
        print("Valuuttakurssit file does not exist.")
        return

    return args.input, args.valuuttakurssit

def find_currency_rate(date: str, valuuttakurssit_df: pd.DataFrame) -> float:
    # Convert the date to the format in the valuuttakurssit file
    # 02/29/2024 -> 29.2.2024
    date_split = date.split("/")
    date_corrected = f"{int(date_split[1])}.{int(date_split[0])}.{date_split[2]}"
    date_tobefound = datetime.strptime(date_corrected, "%d.%m.%Y") # Same date but in datetime format

    # Find the closest currency rate from the valuuttakurssit file.
    # The dates are from the newest to the oldest. So we need to find the first date that is smaller than the given date.
    previous_date = None
    for i in range(len(valuuttakurssit_df)):
        # Check if the date is exactly the same.
        if valuuttakurssit_df["title"].values[i] == date_corrected:
            currency_rate = valuuttakurssit_df.loc[i, "value"]
            break
        # Convert the date from the csv to datetime format.
        date_valuuttakurssit = datetime.strptime(valuuttakurssit_df["title"].values[i], "%d.%m.%Y")
        # Compare dates.
        if date_valuuttakurssit < date_tobefound:
            # Found the first date that is smaller than the given date.
            # Check if the previous date is closer to the given date.
            if previous_date is not None:
                diff = date_tobefound - date_valuuttakurssit
                diff_previous = previous_date - date_tobefound
                print("Didn't find the exact date for %s. Closest dates are %s and %s." % (date_corrected, valuuttakurssit_df["title"].values[i], valuuttakurssit_df["title"].values[i-1]), end=" ")

                if diff_previous < diff:
                    currency_rate = valuuttakurssit_df.loc[i-1, "value"]
                    print("%s is closer." % valuuttakurssit_df["title"].values[i-1])
                else:
                    currency_rate = valuuttakurssit_df.loc[i, "value"]
                    print("%s is closer." % valuuttakurssit_df["title"].values[i])

                break
        previous_date = date_valuuttakurssit

    return float(currency_rate.replace(",", "."))

def query_exchange_rate() -> float:
    result = None
    while result is None:
        exchange_rate = input()

        # Check if the exchange rate is in the format "1,2345" or "1.2345"
        if "," in exchange_rate:
            exchange_rate = exchange_rate.replace(",", ".")
        # Check if the dollar sign is included
        if "$" in exchange_rate:
            exchange_rate = exchange_rate.replace("$", "")
        try:
            result = float(exchange_rate)
        except ValueError:
            print("Invalid exchange rate. Please enter the exchange rate in the format 1.2345 or 1,2345.")
    return result

class sell_event_details:
    def __init__(self, date_sold):
        self.qty = 0
        self.symbol = ""

        # sell details
        self.date_sold = date_sold
        self.gain_loss = 0
        self.total_proceeds_usd = 0 
        self.currency_rate_sold = 0
        
        # purchase details
        self.date_acquired = None
        self.total_cost_basis_usd = 0
        self.currency_rate_acquired = 0

    def vest_date_fmv_eur(self):
        return self.total_cost_basis_usd / self.currency_rate_acquired

    def total_proceeds_eur(self):
        return self.total_proceeds_usd / self.currency_rate_sold
    
    def gain_loss_eur(self):
        return self.total_proceeds_eur() - self.vest_date_fmv_eur()

def create_html_report(sell_events: list[sell_event_details]):

    # Use comma as the decimal separator in the output, because this is the format used in the Finnish tax return.
    import locale
    locale.setlocale(locale.LC_ALL, 'fi_FI.UTF-8')

    # Create an HTML report for the sell events.
    # Create a table
    html = []
    html.append("<html><body>")
    html.append("<h1>E*TRADE-osakkeiden luovutukset vuonna 2025</h1>")
    html.append("<p>Vuonna 2025 luovutetut osakkeet E*TRADE-palvelussa.</p>")
    html.append("<table border='1' cellspacing='0' cellpadding='2'>")
    html.append("<tr>")
    html.append("<th>#</th>") #index
    html.append("<th>Arvopaperin<br>nimi</th>")
    html.append("<th>Arvopaperin<br>laji*</th>")
    html.append("<th>Lukumäärä</th>")
    html.append("<th>Hankintapäivä</th>")
    html.append("<th>Hankintahinta [EUR]</th>")
    html.append("<th>Luovutuspäivä</th>")
    html.append("<th>Luovutushinta [EUR]</th>")
    html.append("<th>Voitto/tappio [EUR]</th>")
    html.append("</tr>")

    total_gain_loss_eur = 0
    total_gain_eur = 0
    total_loss_eur = 0
    total_proceeds_eur = 0
    total_qty = 0
    for idx, sell_event in enumerate(sell_events, 1):
        html.append("<tr>")
        html.append("<td>%d</td>" % idx) # index, starting from 1
        html.append("<td>%s</td>" % sell_event.symbol) # Arvopaperin nimi
        html.append("<td align='center'>51" ) # Arvopaperin laji
        html.append("<td align='right'>%d</td>" % sell_event.qty) # Lukumäärä
        html.append("<td align='right'>%s</td>" % convert_date_to_finnish_format(sell_event.date_acquired)) # Hankintapäivä
        html.append(locale.format_string("<td align='right'>%.2f</td>", sell_event.vest_date_fmv_eur())) # Hankintahinta [EUR]
        html.append("<td align='right'>%s</td>" % convert_date_to_finnish_format(sell_event.date_sold)) # Luovutuspäivä
        html.append(locale.format_string("<td align='right'>%.2f</td>", sell_event.total_proceeds_eur())) # Luovutushinta [EUR]
        html.append(locale.format_string("<td align='right'>%.2f</td>", sell_event.gain_loss_eur())) # Voitto/tappio [EUR]
        html.append("</tr>")

        total_proceeds_eur += sell_event.total_proceeds_eur()
        total_gain_loss_eur += sell_event.gain_loss_eur()
        if sell_event.gain_loss_eur() > 0:
            total_gain_eur += sell_event.gain_loss_eur()
        else:
            total_loss_eur += sell_event.gain_loss_eur()
        total_qty += sell_event.qty

    html.append("</table>")
    html.append("<br>*Arvopaperin laji 51 = Perusosake (ulkomainen)</td>")

    html.append("<br>")
    html.append("<h2>Veroilmoituksen kannalta oleelliset tiedot:</h2>")
    html.append("<table border='0' cellspacing='0' cellpadding='2'>")
    html.append("<td align='right'><b>Luovutushinnat yhteensä:</b></td>")
    html.append(locale.format_string("<td align='right'>%.2f €</td>", total_proceeds_eur))
    html.append("</tr>")
    html.append("<tr>")
    html.append("<td align='right'><b>Luovutusvoitot yhteensä:</b></td>")
    html.append(locale.format_string("<td align='right'>%.2f €</td>", total_gain_eur))
    html.append("</tr>")
    html.append("<tr>")
    html.append("<td align='right'><b>Luovutustappiot yhteensä:</b></td>")
    html.append(locale.format_string("<td align='right'>%.2f €</td>", abs(total_loss_eur)))
    html.append("</tr>")
    html.append("</table>")

    html.append("<br>")
    html.append("<h2>Muut tiedot:</h2>")
    html.append("<table border='0' cellspacing='0' cellpadding='2'>")
    html.append("<tr>")
    html.append("<td align='right'><b>Osakkeet yhteensä:</b></td>")
    html.append(locale.format_string("<td align='right'>%d</td>", total_qty))
    html.append("</tr>")
    html.append("<tr>")
    html.append("<td align='right'><b>Voitto/tappio yhteensä:</b></td>")
    html.append(locale.format_string("<td align='right'>%.2f €</td>", total_gain_loss_eur))
    html.append("</tr>")
    html.append("</table>")
    html.append("</body></html>")

    with open("etrade_luovutukset_2025.html", "w", encoding="utf-8") as f:
        f.write("\n".join(html))
    print("All sell events details saved to etrade_luovutukset_2025.html")


def convert_date_to_finnish_format(date: str) -> str:
    # Convert the date to the format in the valuuttakurssit file
    # 02/29/2024 -> 29.2.2024
    date_split = date.split("/")
    finnish_formatted_date = f"{int(date_split[1])}.{int(date_split[0])}.{date_split[2]}"
    return finnish_formatted_date

def main():
    # check if pandas and openpyxl are installed
    try:
        import pandas as pd
        import openpyxl
    except ImportError:
        print("This script requires pandas and openpyxl. Please install them with 'pip install pandas openpyxl'")
        return

    input_file, valuuttakurssit = parse_args()
    print(input_file, valuuttakurssit)

    # Read the excel file
    df = pd.read_excel(input_file)

    # Get the required columns.
    record_type = df["Record Type"] # This column contains the type of the record. We are interested in "Sell" records.
    qty_a = df["Quantity"]
    symbol_a = df["Symbol"] # This column contains the stock symbol.
    date_acquired_a = df["Date Acquired"] # Note: this is not exactly the same as column "Vest Date FMV". ESPP doesn't have "Vest Date FMV"
    cost_basis_per_share_a = df["Adjusted Cost Basis Per Share"] #  Note: seems to be same as column "Vest Date FMV" for RSU, but ESPP doesn't have "Vest Date FMV"
    adjusted_cost_basis_a = df["Adjusted Cost Basis"]
    date_sold_a = df["Date Sold"]
    total_proceeds_a = df["Total Proceeds"]
    proceeds_per_share_a = df["Proceeds Per Share"]
    adjusted_gain_loss_a = df["Adjusted Gain/Loss"]
    plan_type_a = df["Plan Type"] # RS(U) or ESPP
    # Check if the ESPP exchange rate is in the file. If not, add it.
    if "ESPP Exchange Rate" not in df.columns:
        df["ESPP Exchange Rate"] = pd.NaT
    espp_exchange_rate_a = df["ESPP Exchange Rate"] # This is added by this script to store the exchange rates used when an ESPP was purchased.

    # Verify all the columns are the same length
    assert len(date_acquired_a) == len(qty_a) == len(cost_basis_per_share_a) == len(adjusted_cost_basis_a) \
        == len(date_sold_a) == len(proceeds_per_share_a) == len(total_proceeds_a) == len(adjusted_gain_loss_a) \
        == len(plan_type_a)
    
    assert(record_type[0] == "Summary") # Should be a summary row
    sanity_check_total_qty = qty_a[0] # The total quantity in the summary row should match the sum of the non-rounded (and rounded) quantities in the sell records.
    sanity_check_adjusted_gain_loss_usd = adjusted_gain_loss_a[0] # The total adjusted gain/loss in USD in the summary row should match the sum of the adjusted gain/loss in the sell records.

    assert abs(sanity_check_total_qty - sum(qty_a[1:])) < 0.01, "Total quantity in summary row does not match the sum of the quantities in the sell records."
    # Same but round the quantities in the sell records to nearest integer, because there are some entries with fractional shares.
    assert abs(sanity_check_total_qty - sum(round(qty) for qty in qty_a[1:])) < 0.001, "Total quantity in summary row does not match the rounded sum of the quantities in the sell records."
    assert abs(sanity_check_adjusted_gain_loss_usd - sum(adjusted_gain_loss for adjusted_gain_loss in adjusted_gain_loss_a[1:])) < 0.01, "Total adjusted gain/loss in summary row does not match the sum of the adjusted gain/loss in the sell records."
    
    # Look for ESPP entries. The exchange rate for ESPP purchase should be used from the "ESPP purchase confirmation" document.
    espp_dates = []
    for i in range(len(plan_type_a)):
        if plan_type_a[i] == "ESPP" and espp_exchange_rate_a[i] is pd.NaT:
            espp_dates.append((date_acquired_a[i], i))

    if len(espp_dates) > 0:
        print("ESPP entries found. The EUR->USD exchange rate for ESPP purchase should be used from the \"ESPP purchase confirmation\" document.")
        print("Download the document from etrade and find the exchange rate for the ESPP purchase date.")
        print("At work -> My Account -> Stock Plan Confirmation. Select the wanted year and click 'Apply'.")
        print("Find the ESPP sell orders. Expand the order and click 'View Confirmation of Purchase' under 'Purchase / Grant Details'.")
        print("This should give you a PDF. Look for the 'Average Exchange Rate' under 'Purchase Details' and 'Contributions'.")
        print("This is the exchange rate used by etrade to exchange your EUR to USD before purchasing the ESPP.")

        print("Find exchange rates for the following ESPP purchased on following days:")
        for date in espp_dates:
            print(" - ", date[0])
        
        print("After you have found the exchange rates, you can continue with this script.")
        for date in espp_dates:
            print(" Enter exchange rate for %s: " % date[0], end="")
            exchange_rate = query_exchange_rate()
            print("Using exchange rate %f for ESPP purchased on %s" % (exchange_rate, date[0]))
            # Add the exchange rate to the excel file.
            espp_exchange_rate_a[date[1]] = exchange_rate # date[1] is the index of the espp purchase in the excel file.
            print("------")
        # Save the exchange rates to the file.
        try:
            modified_file_name = input_file.replace(".xlsx", "_modified.xlsx")
            df.to_excel(modified_file_name, index=False)
            print("All exchange rates added. Exchange rates saved to the file: %s" % modified_file_name)
        except Exception as e:
            print("Failed to save the exchange rates to the file.")
            print("Error: ", e)
            print("Continue with the script? (y/n)")
            answer = input()
            if answer.lower() != "y":
                return
        print("---------------------------------------")

    # Read the valuuttakurssit file. Skip first 3 lines because there are some garbage.
    valuuttakurssit_df = pd.read_csv(valuuttakurssit, skiprows=3)

    sell_events = [] # list to store the sell event details.

    for i in range(len(date_acquired_a)):

        date_acquired        = date_acquired_a[i]
        if pd.isna(date_acquired):
            continue

        qty_not_rounded      = qty_a[i] # For some reason there are some entries with fractional shares. 
        qty                  = round(qty_not_rounded) # Round qty to nearest integer. This shouldn't matter, because the total proceeds and gain/loss should be the same regardless of the rounding, but we need to round it for sanity checks and for calculating the gain/loss in EUR.
        cost_basis_per_share = cost_basis_per_share_a[i]
        adjusted_cost_basis  = adjusted_cost_basis_a[i] 
        date_sold            = date_sold_a[i]
        proceeds_per_share   = proceeds_per_share_a[i]
        total_proceeds       = total_proceeds_a[i]
        adjusted_gain_loss   = adjusted_gain_loss_a[i] # for sanity check, should be the same as (total_proceeds - adjusted_cost_basis) or (proceeds_per_share - cost_basis_per_share) * qty

        # Sanity checks
        threshold = 0.01  # 1 cent threshold for sanity checks, because there might be some rounding issues.
        assert abs(adjusted_cost_basis - cost_basis_per_share * qty_not_rounded) < threshold, "Adjusted cost basis does not match cost basis per share * qty for index %d: %f != %f" % (i, adjusted_cost_basis, cost_basis_per_share * qty_not_rounded)
        assert abs(adjusted_gain_loss - (total_proceeds - adjusted_cost_basis)) < threshold, "Adjusted gain/loss does not match total proceeds - adjusted cost basis for index %d: %f != %f" % (i, adjusted_gain_loss, total_proceeds - adjusted_cost_basis)
        assert abs(adjusted_gain_loss - ((proceeds_per_share - cost_basis_per_share) * qty)) < threshold, "Adjusted gain/loss does not match (proceeds per share - cost basis per share) * qty for index %d: %f != %f" % (i, adjusted_gain_loss, (proceeds_per_share - cost_basis_per_share) * qty)

        currency_rate_acquired = find_currency_rate(date_acquired, valuuttakurssit_df)
        vest_date_fmv_eur = qty * cost_basis_per_share / currency_rate_acquired

        currency_rate_sold = find_currency_rate(date_sold, valuuttakurssit_df)
        total_proceeds_eur = qty * proceeds_per_share / currency_rate_sold
        gain_loss = total_proceeds_eur - vest_date_fmv_eur

        # Store values in a sell event details object.
        sell_event = sell_event_details(date_sold)
        sell_event.qty = qty
        sell_event.symbol = symbol_a[i]

        # sell details
        sell_event.gain_loss = gain_loss
        sell_event.total_proceeds_usd = total_proceeds
        sell_event.currency_rate_sold = currency_rate_sold

        # purchase details
        sell_event.date_acquired = date_acquired
        sell_event.total_cost_basis_usd = adjusted_cost_basis
        sell_event.currency_rate_acquired = currency_rate_acquired

        sell_events.append(sell_event)

    create_html_report(sell_events)

if __name__ == "__main__":
    main()