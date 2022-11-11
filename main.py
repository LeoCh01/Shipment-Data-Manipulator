import tkinter as tk
from tkinter import messagebox
import pandas as pd
import numpy as np
from math import ceil
from collections import OrderedDict

items = [
    'Invoice Date',
    'Air Waybill Number',
    'Service Type',
    'Ship Date',
    'Rated Weight Amount',
    'Rated Weight Units',
    'Pieces',
    'SvcPkg Label',
    'Zone Code',
    'Net Transportation Charges',
    'Fuel Surcharge'
]

items_to_zero = [
    'Peak Surcharge',
    'Residential Delivery',
    'Address Correction',
    'Extended Delivery Area',
    'Declared Value Charge',
    'Out of Delivery Area Tier B',
    'Additional Handling Charge - Weight',
    'Additional Handling Charge - Packaging',
    'Additional Handling Charge - Dimensions',
    'Third Party Billing',
    'Saturday Pickup',
    'Canada HST (ON)',
    'Canada GST',
    'Canada HST (NS)',
    'Canada HST (NF)',
    'Canada HST (PE)',
    'Canada HST (NB)',
    'Canada QST',
    'Total'
]

tax = [
    ('Canada HST (ON)', 0.13),
    ('Canada GST', 0.05),
    ('Canada HST (NS/NF)', 0.15),
    ('Canada HST (PE/NB)', 0.15),
    ('Canada QST', 0.1)
]


def on_click():
    try:
        df = pd.read_csv('res/sheet.csv')

        for i in items_to_zero:
            df[i] = 0

        df = df.astype({'Air Waybill Number': str})
        df = set_charge_label(df)
        df['Residential Delivery'] *= 2
        df['ntc old'] = df['Net Transportation Charges']
        df['Additional Handling Charge'] = (
            df['Additional Handling Charge - Weight'] +
            df['Additional Handling Charge - Packaging'] +
            df['Additional Handling Charge - Dimensions']
        )

        df = weight_conv(df)
        df = set_ntc(df)
        df = set_fuel_and_tax(df)

        df = set_tot(df)
        items.append('Total')
        df = df[items]
        df = set_tot_row(df)

        df = df.replace(0, np.nan)
        df = df.dropna(axis=1, how='all')

        df.to_csv('res/new_sheet.csv', index=False)
        messagebox.showinfo('Return', 'your file has been converted.')

    except Exception as e:
        messagebox.showerror('Error', f'Error at line {e.__traceback__.tb_lineno}:\n{e}')


def weight_conv(df):
    df.loc[df['Rated Weight Units'] == 'K', 'Rated Weight Amount'] *= 2.20462
    df['Rated Weight Amount'] = round(df['Rated Weight Amount'], 1)
    df['Rated Weight Units'] = 'L'
    return df


def set_ntc(df):
    ipe = pd.read_csv('res/IPE.csv', header=4)
    ip = pd.read_csv('res/IP.csv', header=1)
    ie = pd.read_csv('res/IE.csv', header=1)
    dom = pd.read_csv('res/DOM.csv', header=1)
    d = {
        'FedEx International Priority': [(ip, ipe), {'Env': (0, 1), 'Pak': (4, 2), 'Pack': (9, 100)}],
        'FedEx Economy': [dom, {'Env': (0, 1), 'Pak': (4, 1), 'Pack': (8, 100)}],
        'FedEx International Economy': [ie, {'Env': (0, 100), 'Pak': (0, 100), 'Pack': (0, 100)}],
        'FedEx Standard Overnight': [dom, {'Env': (0, 1), 'Pak': (4, 1), 'Pack': (8, 100)}],
    }

    no_rate = [''] * len(df)
    ntc = []
    for i in range(len(df)):
        service = df['Service Type'][i]
        zone = df['Zone Code'][i]
        size = ceil(df['Rated Weight Amount'][i])
        label = df['SvcPkg Label'][i]
        t = int(df['POD Time'][i][-4:]) if df['POD Time'][i][-4:] != ' ' else 1030

        if service in d:
            if service == 'FedEx Standard Overnight':
                ntc.append(-3)
                no_rate[i] = 'No Rate'
            elif 'Env' in label or 'Pak' in label or 'Pack' in label:
                ntc.append(float(get_cost(size, d[service], label, zone, t)))
            else:
                ntc.append(-2)
        else:
            ntc.append(-1)

    df['Net Transportation Charges'] = pd.Series(ntc)
    df['Rate'] = pd.Series(no_rate)
    df.loc[df['Rate'] == 'No Rate', 'Net Transportation Charges'] = df['ntc old'] * 1.2
    return df


def get_cost(n, lst, label, zone, t):
    zone = str(int(zone)) if zone.isnumeric() else zone
    f, d = lst

    if isinstance(f, tuple):
        f = f[0] if t >= 1030 else f[1]

    if 'Env' in label and n <= d['Env'][1]:
        return f[zone][d['Env'][0] + ceil(n) - 1]
    if ('Env' in label or 'Pak' in label) and n <= d['Pak'][1]:
        return f[zone][d['Pak'][0] + ceil(n) - 1]
    if n <= d['Pack'][1]:
        return f[zone][d['Pack'][0] + ceil(n) - 1]
    if n <= 300:
        return f[zone][f.index[f['Weight'] == '300.00'].tolist()[0]] * n
    if n <= 500:
        f[zone][f.index[f['Weight'] == '500.00'].tolist()[0]] * n
    if n <= 1000:
        f[zone][f.index[f['Weight'] == '1000.00'].tolist()[0]] * n
    if n <= 2000:
        f[zone][f.index[f['Weight'] == '2000.00'].tolist()[0]] * n
    else:
        return -3


def set_charge_label(df):
    df2 = df.filter(like='Air Waybill Charge')
    df2 = df2.dropna(axis=1, how='all')
    df2 = df2.transpose()
    add_items = OrderedDict()
    for i in range(len(df2) // 2):
        for j in range(len(df2.columns)):
            if df2.iloc[i * 2][j]:
                add_items[df2.iloc[i * 2][j]] = None
                df.at[j, df2.iloc[i * 2][j]] = df2.iloc[i * 2 + 1][j]

    add_items.pop('Transportation Charge')
    add_items.pop('Volume Discount')
    if 'Subtotal' in add_items:
        add_items.pop('Subtotal')
    for i in add_items.keys():
        if i not in items:
            items.append(i)

    # df2.to_csv('test.csv', index=False)
    return df


def set_fuel_and_tax(df):
    df.loc[df.shape[0]] = [0] * len(df.columns)
    z, zz = get_z(df)
    df['Fuel %'] = df['Fuel Surcharge'] / (z + df['ntc old'] + (df['Residential Delivery'] / 2))
    df['Fuel Surcharge'] = round(df['Fuel %'] * (z + df['Net Transportation Charges'] + df['Residential Delivery']), 2)

    t = z + zz + df['Residential Delivery'] + df['Fuel Surcharge'] + df['Net Transportation Charges'] # ?

    df['Canada HST (NS/NF)'] = df['Canada HST (NS)'] + df['Canada HST (NF)']
    df['Canada HST (PE/NB)'] = df['Canada HST (PE)'] + df['Canada HST (NB)']

    for i, j in tax:
        df.loc[df[i] > 0, i] = t * j
    return df


def get_z(df):
    z = (
        df['Peak Surcharge'] +
        df['Extended Delivery Area'] +
        df['Additional Handling Charge'] +
        df['Saturday Pickup'] +
        df['Out of Delivery Area Tier B']
    )
    zz = (
        df['Address Correction'] +
        df['Declared Value Charge'] +
        df['Third Party Billing']

    )
    return z, zz


def set_tot(df):
    for i in items[9:]:
        df[i] = df[i].replace(np.nan, 0)
        df['Total'] += df[i]
    return df


def set_tot_row(df):
    for i, j in zip(items[7:], df.dtypes[7:]):
        if j == 'float64':
            df.at['tot', i] = df[i].sum()
    df.at['tot', 'Invoice Date'] = 'Total:'
    return df


root = tk.Tk()
root.attributes('-topmost', True)
root.title('XCHANGE')
root.geometry('200x100')
root.resizable(width=False, height=False)
text = tk.Text(root)
text.tag_add('sheet.csv', '1.0', '1.4')
text.tag_config('sheet.csv', foreground='blue')
b = tk.Button(root, text='Convert', height=3, width=12, fg='slate blue', font='size 12 bold', command=on_click)
b.place(x=35, y=15)
root.mainloop()
