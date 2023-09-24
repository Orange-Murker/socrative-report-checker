# The directory where the sheets are
DIR_NAME = "Sprint 1"

# Use regular expressions if team members have similar names
team_members = {
    22: ["Name1", "Name2"],
    23: ["Name3", "Name4$"]
}

# ------------------------------

import pandas as pd
import os
import re
import itertools
from collections import defaultdict

red = "\033[0;31m"
green = "\033[0;32m"
yellow = "\033[0;33m"
cyan = "\033[0;36m"
reset = '\033[0m'


member_teams = {}
for team, names in team_members.items():
    for name in names:
        member_teams[name] = team

member_names = list(itertools.chain.from_iterable(team_members.values()))

# Dictionary to store member names and the teams they gave feedback to
gave_feedback = defaultdict(list)

os.chdir(DIR_NAME)

files = filter(lambda x: x.endswith(".xlsx"), os.listdir())

for file in files:
    # Ignore the general feedback file
    if "General" not in file:
        current_team = int(file.split("Team ")[1][:2])
        print(f"\nProcessing feedback for team {current_team}")

        df = pd.read_excel(file, sheet_name="Sheet1")

        feedback_names = df.iloc[:, 0][6:].fillna("")
        feedback_teams = df.iloc[:, 4][6:].fillna("")

        for i, feedback_name in enumerate(feedback_names):
            # Ignore socrative stuff
            if feedback_name and "Report" not in feedback_name and "Scoring" not in feedback_name:
                found = False
                for member_name in member_names:
                    if re.search(member_name, feedback_name, re.IGNORECASE):
                        found = True
                        gave_feedback[member_name].append(current_team)

                        # Check if the person filled out the correct team number
                        correct_number = str(member_teams[member_name])
                        actual_number = str(feedback_teams.iloc[i])
                        if actual_number and correct_number not in actual_number:
                            print(f"{red}{member_name} filled out a wrong team number {actual_number} when they are {correct_number}{reset}")
                if not found:
                    print(f"{yellow}{feedback_name} could not be found{reset}")


for team, names in team_members.items():
    print(f"\n{cyan}Team {team} report:{reset}")
    for name in names:
        feedback = gave_feedback[name]
        colour = reset
        expected_feedback = len(team_members) - 1
        if len(feedback) == expected_feedback:
            colour = green
        elif len(feedback) > expected_feedback:
            colour = yellow
        else:
            colour = red
        print(f"{name:<20}: gave feedback to: {colour}{feedback}{reset}")
