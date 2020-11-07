import json
from datetime import datetime
import os
from os import listdir, path
from os.path import isfile, join, isdir, exists
import gspread
from gspread import utils
from oauth2client.service_account import ServiceAccountCredentials
import sys
import re


def init():
    config_path = join(data_path, "SkillCollectorConfig.json")
    if not exists(config_path):
        end_program("SkillCollectorConfig.json not found")
    else:
        print("SkillCollectorConfig.json loaded")
    with open(config_path, 'r', encoding='utf-8') as data:
        conf: dict = json.load(data)

    # Write default values to the config file if the key don't exists
    conf["sheet_id"] = conf.setdefault("sheet_id", "")
    conf["worksheet_name"] = conf.setdefault("worksheet_name", "")
    conf["date_row_name"] = conf.setdefault("date_row_name", "datum")
    conf["date_format"] = conf.setdefault("date_format", "%d.%m.%Y")
    conf["skill_columns"] = conf.setdefault("skill_columns", "C:D")

    with open(config_path, 'w+', encoding='utf-8') as data:
        json.dump(conf, data, indent=2)
    return conf


class Player:
    def __init__(self, player_name, log_path, date):
        self.player_name = player_name
        self.player_name_print = player_name.capitalize()
        self.date = date
        self.skills = extract_skills(log_path)


def main():
    try:
        print("Extract all skill files...\n")
        players = extract_all_logs(data_path)

        any_found = False
        for player in players:
            print("Latest skills for '" + player.player_name_print + "' on '" + str(player.date) + "' found")

            if player.skills.__len__() == 0:
                print("\tNo skills found inside the file!")
                player.found = False
            else:
                any_found = True

        if not any_found:
            end_program("\nNo skill files found!")

        print("\nConnect to GoogleDocs...")

        scope = ['https://www.googleapis.com/auth/spreadsheets']
        worker = config["service_worker"]
        credentials = ServiceAccountCredentials._from_parsed_json_keyfile(worker, scope)
        client = gspread.authorize(credentials)
        sheet = client.open_by_key(config["sheet_id"]).worksheet(config["worksheet_name"])

        all_values = sheet.get_all_values()

        skill_rows = []
        names = []

        skill_range_indexes = utils.a1_range_to_grid_range(config["skill_columns"])
        range_start_index = skill_range_indexes["startColumnIndex"]
        range_end_index = skill_range_indexes["endColumnIndex"]

        # Preparing skill rows
        for row in all_values:
            for i in range(range_start_index, range_end_index):
                if players[0].skills.__contains__(row[i].lower()) or config["date_row_name"].lower() == row[i].lower():
                    skill_rows.append(row[i].lower())
                    break
            else:
                skill_rows.append("")

        # Preparing names
        for single_name in all_values[4]:
            names.append(single_name.lower())

        # Writing player skills in table
        for player in players:
            if not names.__contains__(player.player_name):
                continue

            print("\nPreparing data...")
            player_index = names.index(player.player_name)
            player_cell = utils.rowcol_to_a1(1, player_index + 1)
            old_values = [row[player_index] for row in all_values]

            new_values = []
            updated_skills = []

            # Set all rows of the player column
            for i in range(len(skill_rows)):
                if skill_rows.__len__() > i and player.skills.__contains__(skill_rows[i].lower()):
                    cur_skill = round(float(player.skills[skill_rows[i].lower()]) * 100) / 100
                    new_values.append([cur_skill])

                    if old_values.__len__() > i:
                        try:
                            if cur_skill - float(old_values[i].replace(",", ".")) > 0:
                                updated_skills.append((skill_rows[i], old_values[i], cur_skill))
                        except Exception as e:
                            updated_skills.append((skill_rows[i], old_values[i], cur_skill))
                    else:
                        updated_skills.append((skill_rows[i], 0, cur_skill))
                elif skill_rows[i].lower() == config["date_row_name"]:
                    new_values.append([player.date.strftime(config["date_format"])])
                else:
                    new_values.append([])

            print("Writing data into the table...")
            sheet.update(player_cell, new_values)

            # Print changed skills
            if updated_skills.__len__() > 0:
                print("Changed skills for '" + player.player_name_print + "':")
                for skill in updated_skills:
                    whitespaces = " " * max(1, 25 - len(skill[0]))
                    print("\t" + skill[0].capitalize() + ": " + whitespaces + to_number(skill[1]) + " -> " + to_number(skill[2]))
                print("\nWriting successful!\n")
            else:
                print("No skills changed for '" + player.player_name_print + "'")
    except Exception as e:
        print(e)
        end_program("", -1)
    end_program()


def to_number(value):
    return str(value).replace(".", ",")


def extract_all_logs(wurm_path) -> list:
    paths = ["gamedata/players", "players", "../gamedata/players", "../players"]
    players = []

    for cur_path in paths:
        full_cur_path = join(wurm_path, cur_path)
        if exists(full_cur_path):
            for player_folder in listdir(full_cur_path):
                full_player_path = join(full_cur_path, player_folder)
                full_dumps_path = join(full_player_path, "dumps")

                if not isdir(full_player_path):
                    continue
                if not exists(full_dumps_path):
                    continue

                times = []

                for file in listdir(full_dumps_path):
                    if not isfile(join(full_dumps_path, file)):
                        continue
                    if bool(re.match("^skills\\.[0-9]{8}\\.[0-9]{4}\\.txt$", file)):
                        times.append((datetime.strptime(file[7:20], "%Y%m%d.%H%M"), file))

                if times.__len__() > 0:
                    times.sort(reverse=True)
                    log_path = join(full_cur_path, player_folder, "dumps", times[0][1]).replace("\\", "/")
                    players.append(Player(player_folder.lower(), log_path, times[0][0]))
    return players


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


if __name__ == '__main__':
    data_path = os.path.dirname(os.path.realpath(__file__))
    config = init()
    main()
