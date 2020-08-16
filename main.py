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


class Player:
    def __init__(self, player_name, log_path, date, found):
        self.player_name = player_name
        self.player_name_print = player_name.capitalize()
        self.log_path = log_path
        self.date = date
        self.found = found
    player_name: ""


def main():
    try:
        print("Search for skill files...\n")

        if config.__contains__("players"):
            player_names = config["players"]
        else:
            player_names = None

        if player_names is None or player_names.__len__() == 0:
            players = [search_for_newest_log(data_path)]
        else:
            players = [search_for_newest_log(data_path, player_name) for player_name in player_names]

        any_found = False
        for player in players:
            if player.found:
                print("Latest skills for '" + player.player_name_print + "' on '" + str(player.date) + "' found")
                player.skills = extract_skills(player.log_path)

                if player.skills.__len__() == 0:
                    print("\tNo skills found inside the file!")
                    player.found = False
                else:
                    any_found = True
            elif player.player_name is not None:
                print("No skill file found for '" + player.player_name_print + "'")

        if not any_found:
            end_program("\nNo skill files found!")

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

        for player in players:
            if not player.found:
                continue
            if not names.__contains__(player.player_name):
                print("Player '" + player.player_name_print + " not found inside the table! (Row 5)")
                continue

            print("\nPreparing data...")
            player_index = names.index(player.player_name)
            player_cell = utils.rowcol_to_a1(1, player_index + 1)
            old_values = [row[player_index] for row in all_values]

            new_values = []
            updated_skills = []

            for i in range(len(skill_names)):
                if skill_names.__len__() > i and player.skills.__contains__(skill_names[i].lower()):
                    cur_skill = round(float(player.skills[skill_names[i].lower()]) * 100) / 100
                    new_values.append([cur_skill])

                    if old_values.__len__() > i:
                        try:
                            if cur_skill - float(old_values[i].replace(",", ".")) > 0:
                                updated_skills.append((skill_names[i], old_values[i], cur_skill))
                        except Exception as e:
                            updated_skills.append((skill_names[i], old_values[i], cur_skill))
                    else:
                        updated_skills.append((skill_names[i], 0, cur_skill))
                else:
                    new_values.append([])

            print("Write data into the table...")
            sheet.update(player_cell, new_values)

            if updated_skills.__len__() > 0:
                print("Changed skills for '" + player.player_name_print + "':")
                for skill in updated_skills:
                    whitespaces = " " * max(1, 25 - len(skill[0]))
                    print("\t" + skill[0] + ": " + whitespaces + to_number(skill[1]) + " -> " + to_number(skill[2]))
            else:
                print("No skills changed for '" + player.player_name_print + "':")
            print("\nWriting successful!\n")
    except Exception as e:
        print(e)
        end_program("", -1)
    end_program()


def to_number(value):
    return str(value).replace(".", ",")


def search_for_newest_log(wurm_path, player_search=None) -> Player:
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
    if player_name != "":
        return Player(player_name.lower(), log_path, player_log[0], True)
    else:
        return Player(player_search, "", "", False)


def extract_skills(skill_path):
    with open(skill_path, 'r') as data:
        lines = data.read().split('\n')
        found_skills = {}

        for line in lines:
            skill_line = line.strip().replace(':', '').rsplit(' ', 3)
            if skill_line.__len__() == 4 and skill_line[0] != 'Skills':
                skill = skill_line[0].lower()
                value = max(skill_line[1], skill_line[2], skill_line[3])
                if not found_skills.__contains__(skill):
                    found_skills[skill] = value
                else:
                    found_skills[skill] = max(value, found_skills[skill])
        return found_skills


def end_program(reason="", exit_code=0):
    if reason != "":
        print(reason)
    print("\n\nPress enter to quit...")
    input()
    sys.exit(exit_code)


config = init()
main()
