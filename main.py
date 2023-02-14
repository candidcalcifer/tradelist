import pandas as pd
import os
from datetime import *
import matplotlib.pyplot as plt


tickers = ["ALTA", "ASIA", "BCMFA", "BNG", "BRCOL", "CADEPO", "CANHOU", "CANLIL", "CANMFA", "CITGUE", "CITSUD", "CMHC",
           "CPPIBC", "DURHAM", "EDC", "EIB", "FINQUE", "FNFACA", "HALTON", "HAMCTY", "IADB", "IBRD", "IFC", "KBN",
           "KFW", "KOMMUN", "LONDON", "LONGUL", "MNTRL", "MP", "NBMFC", "NBRNS", "NF", "NFAB", "NIAGMU", "NORTHW", "NS",
           "NSMFC", "OMERFT", "ONT", "ONTTFTM", "OPBFIN", "OTTAWA", "PEEL", "PRINCE", "PSPCAP", "Q", "QC", "QHEL",
           "REGINA", "RENTEN", "SCDA", "SOUCOA", "STJOHNS", "TRNT", "UOFTOR", "VANC", "VILSHE", "WINNPG", "WTRLOO",
           "YORKMY", "YUKDEV"]
provs = ["BRCOL", "ALTA", "MP", "SCDA", "NF", "NBRNS", "PRINCE", "NS"]
brokers = ["SHORCAN", "FREEDOM", "TULLETT PREBON"]


def bias(df: pd.DataFrame, name: str):
    buy, sell = volumes(df)

    bar_width = 0.9
    r1 = [1, 4, 7, 10, 13]
    r2 = [2, 5, 8, 11, 14]

    plt.bar(r1, sell, width=bar_width, color=[(36/255, 64/255, 98/255)], label="Investor Buy")
    plt.bar(r2, buy, width=bar_width, color=[(220/255, 230/255, 241/255)], label="Investor Sell")
    plt.legend()
    plt.xticks([1.5, 4.5, 7.5, 10.5, 13.5], ["0-5Y", "5-10Y", "10-20Y", "20-28Y", "28Y+"])
    plt.xlabel("Maturity")
    plt.ylabel("Qty (millions)")
    plt.title(name + "Investor Trade Bias")

    plt.subplots_adjust(bottom=0.2)
    plt.grid(axis="y", color="grey", linestyle="dashed")

    ax = plt.gca()
    ax.set_ylim([0, None])

    plt.savefig(name + ".png")
    plt.clf()


def volumes(df: pd.DataFrame):
    buy = [0, 0, 0, 0, 0]
    sell = [0, 0, 0, 0, 0]
    current = datetime.now().year

    for index, row in df.iterrows():
        if row["Verb"] == "Buy":
            if abs(row["Maturity"].year - current) <= 5:
                buy[0] += row["Quantity"]
            elif abs(row["Maturity"].year - current) <= 10:
                buy[1] += row["Quantity"]
            elif abs(row["Maturity"].year - current) <= 20:
                buy[2] += row["Quantity"]
            elif abs(row["Maturity"].year - current) <= 28:
                buy[3] += row["Quantity"]
            else:
                buy[4] += row["Quantity"]
        if row["Verb"] == "Sell":
            if abs(row["Maturity"].year - current) <= 5:
                sell[0] += row["Quantity"]
            elif abs(row["Maturity"].year - current) <= 10:
                sell[1] += row["Quantity"]
            elif abs(row["Maturity"].year - current) <= 20:
                sell[2] += row["Quantity"]
            elif abs(row["Maturity"].year - current) <= 28:
                sell[3] += row["Quantity"]
            else:
                sell[4] += row["Quantity"]

    return buy, sell


def get_data():
    raw_data = pd.read_excel("Trade Data.xls", sheet_name="Sheet1")

    raw_data.drop(columns=["Id", "CreateTime", "Time", "Amt_Out", "Counterparty", "Count", "RFQ Count", "Lockout",
                           "Source", "NumDealer", "Segment", "Product Group", "Asset Group", "Filter"], inplace=True)

    raw_data.query('Ticker in @tickers', inplace=True)
    raw_data.reset_index(inplace=True)
    raw_data.drop(columns=["index"], inplace=True)

    for i in range(len(raw_data)):
        security = raw_data.at[i, "Security"].split()
        security = security[0] + " " + security[1] + " " + raw_data.at[i, "Maturity"].strftime("%m/%d/%Y")
        raw_data.at[i, "Security"] = security

        raw_data.loc[i, "Quantity"] = round(abs(float(raw_data.loc[i, "Quantity"])))

    return raw_data


def split(trade_data: pd.DataFrame):
    ticker_names = list(set(trade_data["Ticker"]))

    split_directory = {}

    for ticker in ticker_names:
        filtered = trade_data[trade_data["Ticker"] == ticker]
        split_directory[ticker] = filtered

    try:
        split_directory["Q"] = pd.concat([split_directory["Q"], split_directory["QHEL"]])
        split_directory["Q"] = sort_qty(split_directory["Q"])
        del split_directory["QHEL"]
    except KeyError:
        pass

    return split_directory


def sort_qty(df: pd.DataFrame):
    df = df.sort_values(by="Quantity", ascending=False)
    return df


def table(ori: pd.DataFrame):
    ori = ori.loc[ori["Qty"] > 5]
    ori.reset_index(inplace=True)
    ori.drop(columns=["index", "DefaultPrice", "Price", "Cover", "Date", "Quantity", "CounterpartyPrice"], inplace=True)
    return ori


def order(td: dict):
    o = ["CANHOU", "ONT", "Q"]
    new_td = pd.DataFrame()
    misc = pd.Dataframe()

    for ticker in order:
        try:
            tic = td[ticker]
            new_td = pd.concat([new_td, tic])
        except KeyError:
            pass

    remaining_tickers = list(set(td.keys()) - set(o))

    for ticker in remaining_tickers:
        tic = td[ticker]
        misc = pd.concat([misc, tic])

    misc = sort_qty(misc)
    new_td = pd.concat([new_td, misc])
    new_td.drop_duplicates()

    return new_td


def export(d: pd.DataFrame, t: pd.DataFrame, vol: list):
    now = date.today().strftime("%m/%d/%Y")
    vol = pd.DataFrame([vol])

    with pd.ExcelWriter(now + ".xlsx", engine="xlsxwriter") as writer:
        d.to_excel(writer, sheet_name="sorted", index=False)
        t.to_excel(writer, sheet_name="raw", index=False)
        vol.to_excel(writer, sheet_name="vol", index=False)

    workbook = writer.book
    workbook.filename = now + ".xlsm"
    workbook.add_vba_project("./vbaProject.bin")
    workbook.get_worksheet_by_name("sorted").insert_button("J2", {"macro": "report", "caption": "report", "width": 80,
                                                                  "height": 30})
    workbook.get_worksheet_by_name("sorted").insert_button("J5", {"macro": "pdf", "caption": "pdf", "width": 80,
                                                                  "height": 30})

    writer.close()
    os.remove(now + ".xlsx")

    return None


def vols(d: dict):
    v = []
    tickers = ["CANHOU", "ONT", "Q"]

    for a in tickers:
        b = 0
        tb = 0
        s = 0
        ts = 0
        for index, row in d[a].iterrows():
            if row["Verb"] == "Buy":
                tb += row["Qty"]
                if "Done" in row["Status"]:
                    b += row["Qty"]
            elif row["Verb"] == "Sell":
                ts += row["Qty"]
                if "Done" in row["Status"]:
                    s += row["Qty"]
        v.append(str(round(b)) + "/" + str(round(tb)))
        v.append(str(round(s)) + "/" + str(round(ts)))

    b = 0
    tb = 0
    s = 0
    ts = 0
    for c in d.keys():
        if c not in tickers:
            for index, row in d[c].iterrows():
                if row["Verb"] == "Buy":
                    tb += row["Qty"]
                    if "Done" in row["Status"]:
                        b += row["Qty"]
                elif row["Verb"] == "Sell":
                    ts += row["Qty"]
                    if "Done" in row["Status"]:
                        s += row["Qty"]

    v.append(str(round(b)) + "/" + str(round(tb)))
    v.append(str(round(s)) + "/" + str(round(ts)))
    return v


def risk(df: pd.DataFrame):
    return df


if __name__ == '__main__':
    data = get_data()
    data = risk(data)

    sorted_total = split(sort_qty(data))
    for_display = table(order(sorted_total))
    volume = vols(sorted_total)

    export(for_display, data, volume)
    bias(sorted_total["CANHOU"], "CMB")
    bias(sorted_total["ONT"], "Ontario")
    bias(sorted_total["Q"], "Quebec")

