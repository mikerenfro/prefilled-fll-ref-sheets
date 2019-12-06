#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sun Feb 10 16:04:11 2019

@author: renfro
"""

import fll_sheet_utils

timeformat = '%-I:%M %p'  # AM/PM with no leading zeros for hours
# Championship
#(buildingip, buildingrd, buildingcv) = ('Daniels Hall', 'Gym Basement',
#                                        'Henderson Hall')
#column_types = {'Pit': str, 'Team': int, 'Team Name': str,
#                'Time1': float, 'Table1': str,
#                'Time2': float, 'Table2': str,
#                'Time3': float, 'Table3': str,
#                'TimeIP': float, 'PlaceIP': str,
#                'TimeRD': float, 'PlaceRD': str,
#                'TimeCV': float, 'PlaceCV': str}

# Qualifier
(buildingip, buildingrd, buildingcv) = (None, None, None)
column_types = {'Pit': str, 'Team': int, 'Team Name': str,
                'TimeIP': float, 'PlaceIP': str,
                'TimeRD': float, 'PlaceRD': str,
                'TimeCV': float, 'PlaceCV': str,
                'TimePractice': float, 'TablePractice': str,
                'Time1': float, 'Table1': str,
                'Time2': float, 'Table2': str,
                'Time3': float, 'Table3': str}

Debug = False
if Debug:
    # schedule = 'fll-total-schedule.csv'  # Championship
    schedule = 'summertown-total-schedule.csv'  # Qualifier
    rg_template_path = 'city-shaper-score-sheet.pdf'
    rubric_template_path = 'first-lego-league-rubrics.pdf'
    team_template_path = 'summertown-team-template.docx'
    date = '2019-12-07'
else:
    # Non-debugging paths:
    (schedule, rg_template_path,
     rubric_template_path, team_template_path,
     date) = fll_sheet_utils.parse_args()

data = fll_sheet_utils.read_csv(schedule, column_types)
fll_sheet_utils.make_referee_sheets(data, rg_template_path)
fll_sheet_utils.make_judge_sheets(data, rubric_template_path)
fll_sheet_utils.make_team_sheets(data, team_template_path, timeformat,
                                 buildingip, buildingrd, buildingcv, date)
