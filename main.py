from datetime import datetime
import os
from os import listdir, path
from os.path import isfile, join, isdir, exists
import gspread
from gspread import utils
from oauth2client.service_account import ServiceAccountCredentials
import sys
from configparser import ConfigParser

print("Starte", end="")
bundle_dir = getattr(sys, '_MEIPASS', path.abspath(path.dirname(__file__)))
CONFIG = ConfigParser()
CONFIG.read(os.path.join(bundle_dir, "config.ini"))
print("...")


def main():
    try:
        data_path = os.path.dirname(os.path.realpath(__file__))
        # Überschreibe den players_path mit den Installationsordner um lokal testen zu können
        # data_path = "%/Wurm Online"

        print("Suche nach Dateien...\n")
        player, log_path, date = search_for_newest_log(data_path + "/gamedata/players")

        if player == "":
            player, log_path, date = search_for_newest_log(data_path + "/players")

        if player == "":
            end_program("Keine Skilldateien gefunden!")

        print("Neuste Skills von '" + player + "' am '" + str(date) + "' gefunden")
        skills = extract_skills(log_path)

        if skills.__len__() == 0:
            end_program("Keine Skills in der Datei gefunden!")

        print("\nVerbinde zu GoogleDocs...")

        scope = ['https://www.googleapis.com/auth/spreadsheets']
        service_account_path = path.join(bundle_dir, 'service_account.json')
        credentials = ServiceAccountCredentials.from_json_keyfile_name(service_account_path, scope)
        client = gspread.authorize(credentials)
        sheet = client.open_by_key(CONFIG['GoogleDocs']["sheet_id"]).worksheet("Fähigkeiten")

        all_values = sheet.get_all_values()

        skill_names = [row[3] if row[3] > row[2] else row[2] for row in all_values]
        names = all_values[4]

        if not names.__contains__(player):
            end_program("Spieler nicht in der Tabelle gefunden!")

        print("Bereite Daten vor...")
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
            elif old_values.__len__() > i:
                new_values.append([old_values[i]])
            else:
                new_values.append([""])

        print("Schreibe Daten in die Tabelle...")
        sheet.update(player_cell, new_values)
    except Exception as e:
        print(e)
        end_program("", -1)
    else:
        if updated_skills.__len__() > 0:
            print("Geänderte Skills:")
            for skill in updated_skills:
                whitespaces = " " * max(1, 25 - len(skill[0]))
                print("\t" + skill[0] + ": " + whitespaces + to_number(skill[1]) + " -> " + to_number(skill[2]))
        else:
            print("Keine geänderten skills")
        print("\nSchreiben erfolgreich!")
        end_program()


def to_number(value):
    return str(value).replace(".", ",")


def search_for_newest_log(players_path):
    player_name = ""
    player_log = (datetime.min, "")

    if exists(players_path):
        for player_folder in listdir(players_path):
            full_player_path = join(players_path, player_folder)
            full_dumps_path = join(full_player_path, "dumps")

            if not isdir(full_player_path):
                continue
            if not exists(full_dumps_path):
                continue
            logs = [log.replace("skills.", "").replace(".txt", "") for log in listdir(full_dumps_path)
                    if (isfile(join(full_dumps_path, log)) and log.startswith("skills"))]
            times = [(datetime.strptime(log, "%Y%m%d.%H%M"), log) for log in logs]
            times.sort(reverse=True)
            if times.__len__() > 0 and times[0][0] > player_log[0]:
                player_name = player_folder
                player_log = times[0]

    return player_name, join(players_path, player_name, "dumps", "skills." + player_log[1] + ".txt"), player_log[0]


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
    print("\n\nDrücke eine Taste zum beenden...")
    input()
    sys.exit(exit_code)


main()

