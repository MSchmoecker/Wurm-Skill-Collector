from datetime import datetime
import os
from os import listdir, path
from os.path import isfile, join, isdir, exists
import gspread
from gspread import utils
from oauth2client.service_account import ServiceAccountCredentials
import sys
from configparser import ConfigParser


bundle_dir = getattr(sys, '_MEIPASS', path.abspath(path.dirname(__file__)))
CONFIG = ConfigParser()
CONFIG.read(os.path.join(bundle_dir, "config.ini"))


def main():
    try:
        players_path = os.path.dirname(os.path.realpath(__file__)) + "/gamedata/players"
        # Überschreibe den players_path mit den Installationsordner um lokal testen zu können
        # players_path = '%/Wurm Online/gamedata/players

        print("Suche nach Dateien...\n")
        player, log_path, date = search_for_newest_log(players_path)

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

        for i in range(len(skill_names)):
            if skill_names.__len__() > i and skills.__contains__(skill_names[i].lower()):
                new_values.append([round(float(skills[skill_names[i].lower()]) * 100) / 100])
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
        print("Schreiben erfolgreich!")
        end_program()


def search_for_newest_log(players_path):
    player_name = ""
    player_log = (datetime.min, "")

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
        print(reason, 'red')
    print("\n\nDrücke eine Taste zum beenden...")
    input()
    sys.exit(exit_code)


main()

