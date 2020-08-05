import json
from datetime import datetime
import os
from os import listdir, path
from os.path import isfile, join, isdir, exists
import gspread
from gspread import utils
from oauth2client.service_account import ServiceAccountCredentials
import sys

data_path = os.path.dirname(os.path.realpath(__file__))


def init():
    config_path = join(data_path, "SkillCollectorConfig.json")
    if not exists(config_path):
        end_program("SkillCollectorConfig.json not found")
    else:
        print("SkillCollectorConfig.json loaded")
    with open(config_path, 'r', encoding='utf-8') as data:
        return json.load(data)


def main():
    try:
        print("Search for skill files..\n")
        player, log_path, date = search_for_newest_log(data_path)

        if player == "":
            end_program("No skill files found!")

        print("Newest skills for '" + player.capitalize() + "' at '" + str(date) + "' found")
        skills = extract_skills(log_path)

        if skills.__len__() == 0:
            end_program("No skills found inside the file!")

        print("\nConnect to GoogleDocs...")

        scope = ['https://www.googleapis.com/auth/spreadsheets']
        worker = config["service_worker"]
        credentials = ServiceAccountCredentials._from_parsed_json_keyfile(worker, scope)
        client = gspread.authorize(credentials)
        sheet = client.open_by_key(config["sheet_id"]).worksheet(config["worksheet_name"])

        all_values = sheet.get_all_values()

        skill_names = [row[3] if row[3] > row[2] else row[2] for row in all_values]
        names = []
        for single_name in all_values[4]:
            names.append(single_name.lower())

        if not names.__contains__(player):
            end_program("Player name not found inside the table! (Row 5)")

        print("Preparing data...")
        player_index = names.index(player)
        player_cell = utils.rowcol_to_a1(1, player_index + 1)
        old_values = [row[player_index] for row in all_values]

        new_values = []
        updated_skills = []

        for i in range(len(skill_names)):
            if skill_names.__len__() > i and skills.__contains__(skill_names[i].lower()):
                cur_skill = round(float(skills[skill_names[i].lower()]) * 100) / 100
                new_values.append([cur_skill])

                if old_values.__len__() > i:
                    try:
                        if cur_skill - float(old_values[i].replace(",", ".")) != 0:
                            updated_skills.append((skill_names[i], old_values[i], cur_skill))
                    except Exception as e:
                        updated_skills.append((skill_names[i], old_values[i], cur_skill))
                else:
                    updated_skills.append((skill_names[i], 0, cur_skill))
            else:
                new_values.append([])

        print("Write data in the table...")
        sheet.update(player_cell, new_values)
    except Exception as e:
        print(e)
        end_program("", -1)
    else:
        if updated_skills.__len__() > 0:
            print("Changed skills:")
            for skill in updated_skills:
                whitespaces = " " * max(1, 25 - len(skill[0]))
                print("\t" + skill[0] + ": " + whitespaces + to_number(skill[1]) + " -> " + to_number(skill[2]))
        else:
            print("No skills changed")
        print("\nWriting successful!")
        end_program()


def to_number(value):
    return str(value).replace(".", ",")


def search_for_newest_log(wurm_path, player_search=None) -> (str, str, datetime):
    paths = ["gamedata/players", "players", "../gamedata/players", "../players"]
    player_name = ""
    player_log = (datetime.min, "")
    player_path = ""

    for cur_path in paths:
        full_cur_path = join(wurm_path, cur_path)
        if exists(full_cur_path):
            for player_folder in listdir(full_cur_path):
                full_player_path = join(full_cur_path, player_folder)
                full_dumps_path = join(full_player_path, "dumps")

                if player_search is not None and player_folder != player_search:
                    continue
                if not isdir(full_player_path):
                    continue
                if not exists(full_dumps_path):
                    continue
                logs = [log.replace("skills.", "").replace(".txt", "") for log in listdir(full_dumps_path)
                        if (isfile(join(full_dumps_path, log)) and log.startswith("skills"))]
                times = []
                for log in logs:
                    try:
                        times.append((datetime.strptime(log.strip(), "%Y%m%d.%H%M"), log))
                    except Exception as e:
                        pass
                times.sort(reverse=True)
                if times.__len__() > 0 and times[0][0] > player_log[0]:
                    player_name = player_folder
                    player_log = times[0]
                    player_path = full_cur_path

    log_path = join(player_path, player_name, "dumps", "skills." + player_log[1] + ".txt").replace("\\", "/")
    return player_name.lower(), log_path, player_log[0]


def extract_skills(skill_path):
    with open(skill_path, 'r') as data:
        lines = data.read().split('\n')
        found_skills = {}

        for line in lines:
            skill_line = line.strip().replace(':', '').rsplit(' ', 3)
            if skill_line.__len__() == 4 and skill_line[0] != 'Skills':
                found_skills[skill_line[0].lower()] = max(skill_line[1], skill_line[2], skill_line[3])
        return found_skills


def end_program(reason="", exit_code=0):
    if reason != "":
        print(reason)
    print("\n\nPress enter to quit...")
    input()
    sys.exit(exit_code)


config = init()
main()
