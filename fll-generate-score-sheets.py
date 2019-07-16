#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 10 16:04:11 2019

@author: renfro
"""

import io
import pdfrw
from reportlab.pdfgen import canvas
import pandas as pd


def get_overlay_canvas(rg_round, team, table) -> io.BytesIO:
    data = io.BytesIO()
    pdf = canvas.Canvas(data)
    pdf.drawString(x=90, y=550, text=str(rg_round))
    pdf.drawString(x=90, y=575, text=str(team))
    pdf.drawString(x=280, y=550, text=table)
    pdf.save()
    data.seek(0)
    # print("Wrote overlay: round {0}, {1}, {2}".format(rg_round, team, table))
    return data


def merge(overlay_canvas: io.BytesIO, template_path: str, page) -> io.BytesIO:
    template_pdf = pdfrw.PdfReader(template_path)
    overlay_pdf = pdfrw.PdfReader(overlay_canvas)
    data = overlay_pdf.pages[0]
    for page in [template_pdf.pages[page], ]:
        overlay = pdfrw.PageMerge().add(data)[0]
        pdfrw.PageMerge(page).add(overlay).render()
    form = io.BytesIO()
    pdfrw.PdfWriter().write(form, template_pdf)
    form.seek(0)
    return form


def save(form: io.BytesIO, filename: str):
    with open(filename, 'wb') as f:
        f.write(form.read())


def which_round(row, table):
    if row['Table1'] == table:
        return 1
    elif row['Table2'] == table:
        return 2
    elif row['Table3'] == table:
        return 3
    else:
        return None


def which_time(row):
    if row['Round'] == 1:
        return row['Time1']
    elif row['Round'] == 2:
        return row['Time2']
    elif row['Round'] == 3:
        return row['Time3']
    else:
        return None


data = pd.read_table('fll-robot-game-schedule.csv', sep=',')
template_path = './into-orbit-score-sheet.pdf'
table_colors = ['Gold', 'Purple']
table_numbers = [1, 2, 3, 4]
# table_colors = ['Gold', ]
# table_numbers = [1, ]

pd.options.mode.chained_assignment = None

for color in table_colors:
    for number in table_numbers:
        writer = pdfrw.PdfWriter()
        template_pdf = pdfrw.PdfReader(template_path)
        table = '{0} {1}'.format(color, number)
        print(table)
        # For each 8 tables and 48 teams, this probably comes down to:
        # 3x Round 1, 3x Round 2, 3x Round 1, 3x Round 2, 6x Round 3, but this
        # method should be generic for any 3-round tournament setup
        teams_on_table = data[(data['Table1'] == table) |
                              (data['Table2'] == table) |
                              (data['Table3'] == table)]
        teams_on_table['Round'] = teams_on_table.apply(
                lambda row: which_round(row, table), axis=1)
        teams_on_table['Time'] = teams_on_table.apply(
                lambda row: which_time(row), axis=1)
        # Make a bunch of empty sheets for the merge
        for i in range(len(teams_on_table)):
            writer.addpages(template_pdf.pages)
        writer.write('_{0}.pdf'.format(table))

        # This is a bit heavy on the reading/writing, but at 144 total pages,
        # it's not too bad so far.
        page_number = 0
        for index, row in teams_on_table.sort_values('Time').iterrows():
            base_pdf = pdfrw.PdfReader('_{0}.pdf'.format(table))
            canvas_data = get_overlay_canvas(row['Round'], row['Team'], table)
            form = merge(canvas_data, template_path='_{0}.pdf'.format(table),
                         page=page_number)
            page_number = page_number + 1
            save(form, filename='_{0}.pdf'.format(table))
