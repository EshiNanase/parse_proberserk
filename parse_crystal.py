import os
import json
import traceback
from collections import defaultdict

import pandas as pd
import numpy as np
from itertools import permutations
import pathlib

from utils import get_player_to_city_mapper, clean_player, get_city_from_mapper, download_public, extract_id, download_form_data, get_deck_list, get_deck_name, clean_deck_list
from bs4 import BeautifulSoup

TOUR = 4

STANDINGS_FILEPATH = f"data/{TOUR}/standings.htm"
DECKLISTS_FILEPATH = f"deck_lists/{TOUR}/"
DATA_FILEPATH = f"data/{TOUR}/data.json"
FORM_DATA_FILEPATH = "data/form_data.xlsx"


def analyze_turn():

    download_form_data()

    won_turn_amount = 0

    won_after_choosing = 0
    lost_after_choosing = 0

    chosen_first_turn = 0
    wins_after_chosen_first_turn = 0
    losses_after_chosen_first_turn = 0

    chosen_second_turn = 0
    wins_after_chosen_second_turn = 0
    losses_after_chosen_second_turn = 0

    first_turn = 0
    wins_after_first_turn = 0
    second_turn = 0
    wins_after_second_turn = 0

    df = pd.read_excel(FORM_DATA_FILEPATH, sheet_name='Ответы на форму')
    for row in df.values:
        timestamp, gmail, name, tour, won_turn, turn, mulligans, result, *_ = row
        if tour != TOUR:
            continue
        if turn == "Первый":
            first_turn += 1
            if result == "Победа":
                wins_after_first_turn += 1
        else:
            second_turn += 1
            if result == "Победа":
                wins_after_second_turn += 1

        if won_turn == "Да" and result == "Победа":
            won_after_choosing += 1
        elif won_turn == "Да" and result == "Поражение":
            lost_after_choosing += 1

        if won_turn == "Да" and turn == "Первый":
            won_turn_amount += 1
            chosen_first_turn += 1
            if result == "Победа":
                wins_after_chosen_first_turn += 1
            else:
                losses_after_chosen_first_turn += 1

        if won_turn == "Да" and turn == "Второй":
            won_turn_amount += 1
            chosen_second_turn += 1
            if result == "Победа":
                wins_after_chosen_second_turn += 1
            else:
                losses_after_chosen_second_turn += 1

    print(f"{TOUR} тур")
    win_after_first_turn_probability = round(wins_after_first_turn / first_turn * 100, 2)
    print("---")
    print(f"Частота победы 1-ым ходом: {win_after_first_turn_probability}%")

    win_after_second_turn_probability = round(wins_after_second_turn / second_turn * 100, 2)
    print(f"Частота победы 2-ым ходом: {win_after_second_turn_probability}%")

    chosen_first_turn_probability = round(chosen_first_turn / won_turn_amount * 100, 2)
    win_after_chosen_first_turn_probability = round(wins_after_chosen_first_turn / chosen_first_turn * 100, 2)
    loss_after_chosen_first_turn_probability = round(losses_after_chosen_first_turn / chosen_first_turn * 100, 2)
    print("---")
    print(f"Частота выбора 1-го хода: {chosen_first_turn_probability}%")
    print(f"Частота победы, выбрав 1-ый ход: {win_after_chosen_first_turn_probability}%")
    print(f"Частота поражения, выбрав 1-ый ход: {loss_after_chosen_first_turn_probability}%")

    chosen_second_turn_probability = round(chosen_second_turn / won_turn_amount * 100, 2)
    win_after_chosen_second_turn_probability = round(wins_after_chosen_second_turn / chosen_second_turn * 100, 2)
    loss_after_chosen_second_turn_probability = round(losses_after_chosen_second_turn / chosen_second_turn * 100, 2)
    print("---")
    print(f"Частота выбора 2-го хода: {chosen_second_turn_probability}%")
    print(f"Частота победы, выбрав 2-ой ход: {win_after_chosen_second_turn_probability}%")
    print(f"Частота поражения, выбрав 2-ой ход: {loss_after_chosen_second_turn_probability}%")

    win_after_choosing_probability = round(won_after_choosing / won_turn_amount * 100, 2)
    loss_after_choosing_probability = round(lost_after_choosing / won_turn_amount * 100, 2)
    print("---")
    print(f"Частота победы, выбрав ход: {win_after_choosing_probability}%")
    print(f"Частота поражения, выбрав ход: {loss_after_choosing_probability}%")


def analyze_exceeding_copies():

    ignore_cards = ["Цепной пес", "Орк-зомби", "Цепной пес", "Велит", "Помойный крыс"]

    with open(DATA_FILEPATH, "r", encoding="utf-8") as file:
        data = json.loads(file.read())

    for player in data:
        mapper = defaultdict(int)
        for deck in data[player]:
            for card, quantity in data[player][deck].items():
                mapper[card] += quantity
        for card, quantity in mapper.items():
            if quantity > 3 and card not in ignore_cards:
                print(f"'{player}' использовал карту '{card}' {quantity} раз")


def create_standings():

    with open(STANDINGS_FILEPATH, "r", encoding="utf-8") as file:
        data = file.read()

    soup = BeautifulSoup(data, "html.parser")
    standings = {}
    for tr in soup.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 3:
            continue
        index = tds[0].get_text(strip=True)
        player_name = clean_player(tds[1].get_text(strip=True))
        score = tds[2].get_text(strip=True)
        score = score.split(" - ")
        if str(0) in score[-1]:
            score.pop(-1)
        score = "-".join([str(_) for _ in score])
        print(f"Нашел {player_name} со счетом {score} на {index} месте")
        standings[player_name] = {"index": index, "score": score}
    print(f"Создал стендинги с {len(standings)} участниками")
    return standings


def get_downloaded_deck_lists() -> list[str]:

    deck_lists_arr = []
    for path, folders, files in os.walk(DECKLISTS_FILEPATH):
        for file in files:
            deck_lists_arr.append(file)
        for folder in folders:
            for _, __, files in os.walk(folder):
                for file in files:
                    deck_lists_arr.append(file)
    print(f"Уже скачано {len(deck_lists_arr)} деклистов")
    return deck_lists_arr

def download_deck_lists():

    errors = []
    deck_lists = []
    not_found_players = []

    download_form_data()
    mapper = get_player_to_city_mapper()
    standings = create_standings()
    downloaded_deck_lists = get_downloaded_deck_lists()

    os.makedirs(DECKLISTS_FILEPATH, exist_ok=True)
    os.makedirs(DECKLISTS_FILEPATH + "invalid/", exist_ok=True)
    os.makedirs(DECKLISTS_FILEPATH + "uploaded/", exist_ok=True)

    df = pd.read_excel(FORM_DATA_FILEPATH)
    print(f"В DataFrame {len(df)} строк")
    out_dir = pathlib.Path(DECKLISTS_FILEPATH)
    for idx, row in df.iterrows():
        original_player_name = row["Имя + Фамилия"]
        tour = row["Тур"]
        if str(tour) != str(TOUR):
            continue

        try:
            player_name = clean_player(original_player_name)
            player_name_combinations = list(permutations(player_name.split(), 2)) or [[player_name]]
            city = None
            for player_name_combination in player_name_combinations:
                player_new_name = " ".join([_.title() for _ in player_name_combination])
                player_name = clean_player(player_new_name)
                city = get_city_from_mapper(mapper, player_name)
                if city:
                    break

            if not city:
                not_found_players.append(player_name)
                print(f"Не удалось найти игрока {original_player_name}")
                continue

            for deck_ind in [1, 2, 3]:
                deck_url = row[f"{deck_ind} колода"]
                print(f"Игрок {player_name} из города {city} с колодой {deck_url} ({idx})")

                deck_list_name =  f"{deck_ind}_{player_name}_{standings[player_name]['score']}_{standings[player_name]['index']}.txt"
                if deck_list_name in downloaded_deck_lists:
                    print(f"Деклист {deck_list_name} уже скачан")
                    continue

                download_public(extract_id(str(deck_url)), out_dir / deck_list_name)
                deck_lists.append(deck_list_name)
                print(f"Скачал деклист {deck_list_name}")

        except Exception as e:
            print(f"Произошла ошибка: {e}")
            errors.append(f'{original_player_name}: {traceback.format_exc()}')
            continue

    print(f"Было скачано {len(deck_lists)} деклистов")
    print(f"Не найдены следующие игроки: {', '.join(not_found_players)}")
    print(f"Ошибки: {', '.join(errors)}")

def get_banned_decks_excel():

    download_form_data()
    with open("archetypes.txt", "r", encoding="utf-8") as file:
        file_data = file.readlines()
        archetypes = {}
        for line in file_data:
            line = line.replace("\n", "").strip()
            line_arr = line.split(", ")
            archetypes[tuple(line_arr[:-1])] = line_arr[-1]

    mapper = defaultdict(lambda: defaultdict(int))
    unknown_decks = defaultdict(list)
    not_found_decklists = []
    errors = []

    df = pd.read_excel(FORM_DATA_FILEPATH)
    df.replace(np.nan, None, inplace=True)
    print(f"В DataFrame {len(df)} строк")
    for idx, row in df.iterrows():

        try:
            tour = row["Тур"]
            banned_deck_ind = row["Какую колоду вам забанили?"]
            if not banned_deck_ind or tour in {1, 2}:
                continue
            banned_deck = row[f"{int(banned_deck_ind)} колода"]
            download_public(extract_id(str(banned_deck)), pathlib.Path("deck.txt"))
            deck_list = get_deck_list(deck_list_filepath=str(pathlib.Path("deck.txt")))

            if not deck_list:
                print(f"При сборе забаненных колод деклист {deck_list} был пропущен из-за ошибки чтения!")
                continue

            get_deck_name(deck_list)
            deck_list = clean_deck_list(deck_list)
            print(f"Приступил к аналитике забаненных карт в {deck_list}")

            for archetype in archetypes:
                for card in archetype:
                    if card not in deck_list:
                        break
                else:
                    mapper[tour][archetype] += 1
                    print(f"Добавил забаненный деклист к {archetypes[archetype]}")
                    break
            else:
                print(f"Не удалось подобрать архетип для колоды: {banned_deck}")
                unknown_decks[tour].append(banned_deck)
                not_found_decklists.append(banned_deck)

        except Exception as e:
            print(f"Произошла ошибка: {e}")
            errors.append(e)

    print(f"Получилась следующая статистика забаненных архетипов: {mapper}")
    print(f"Не получилось подобрать архетипы для следующих колод: {not_found_decklists}")
    print(f"Произошли следующие ошибки: {errors}")
    with pd.ExcelWriter(pathlib.Path("data", "total_bans.xlsx"), engine="xlsxwriter") as writer:
        for tour in mapper:
            lines = [[archetypes[archetype], mapper[tour][archetype]] for archetype in mapper[tour]]
            lines.append(["Авторская сборка", len(unknown_decks[tour])])
            df = pd.DataFrame(sorted(lines, key=lambda x: x[1], reverse=True), columns=["Архетип", "Кол-во банов"])
            df.to_excel(writer, sheet_name=f"{tour} тур", index=False)
            format_center = writer.book.add_format({'align': 'center'})
            worksheet = writer.sheets[f"{tour} тур"]
            worksheet.set_column('A:A', 60, format_center)
            worksheet.set_column('B:B', 15, format_center)

        total_mapper = defaultdict(int)
        for tour in mapper:
            for archetype in mapper[tour]:
                total_mapper[archetype] += mapper[tour][archetype]

        lines = [[archetypes[archetype], total_mapper[archetype]] for archetype in total_mapper]
        lines.append(["Авторская сборка", sum([unknown_decks[tour] for tour in unknown_decks[tour]])])
        df = pd.DataFrame(sorted(lines, key=lambda x: x[1], reverse=True), columns=["Архетип", "Кол-во банов"])
        df.to_excel(writer, sheet_name="Общая статистика", index=False)
        format_center = writer.book.add_format({'align': 'center'})
        worksheet = writer.sheets["Общая статистика"]
        worksheet.set_column('A:A', 60, format_center)
        worksheet.set_column('B:B', 15, format_center)

        not_found_decklists_df = pd.DataFrame(not_found_decklists, columns=["Неизвестные архетипы"])
        not_found_decklists_df.to_excel(writer, sheet_name="Неизвестные архетипы", index=False)
        worksheet = writer.sheets["Неизвестные архетипы"]
        format_center = writer.book.add_format({'align': 'center'})
        worksheet.set_column('A:AAA', 80, format_center)

def get_picked_decks_excel():

    download_form_data()
    with open("archetypes.txt", "r", encoding="utf-8") as file:
        file_data = file.readlines()
        archetypes = {}
        for line in file_data:
            line = line.replace("\n", "").strip()
            line_arr = line.split(", ")
            archetypes[tuple(line_arr[:-1])] = line_arr[-1]

    mapper = defaultdict(lambda: defaultdict(int))
    unknown_decks = defaultdict(list)
    not_found_decklists = []
    errors = []

    df = pd.read_excel(FORM_DATA_FILEPATH)
    df.replace(np.nan, None, inplace=True)
    print(f"В DataFrame {len(df)} строк")
    for idx, row in df.iterrows():

        try:
            tour = row["Тур"]
            picked_deck_ind = row["Какую колоду вы взяли?"]
            if not picked_deck_ind or tour in {1, 2}:
                continue
            picked_deck = row[f"{int(picked_deck_ind)} колода"]
            download_public(extract_id(str(picked_deck)), pathlib.Path("deck.txt"))
            deck_list = get_deck_list(deck_list_filepath=str(pathlib.Path("deck.txt")))

            if not deck_list:
                print(f"При сборе пикнутых колод деклист {deck_list} был пропущен из-за ошибки чтения!")
                continue

            get_deck_name(deck_list)
            deck_list = clean_deck_list(deck_list)
            print(f"Приступил к аналитике пикнутых карт в {deck_list}")

            for archetype in archetypes:
                for card in archetype:
                    if card not in deck_list:
                        break
                else:
                    mapper[tour][archetypes[archetype]] += 1
                    print(f"Добавил пикнутый деклист к {archetypes[archetype]}")
                    break
            else:
                print(f"Не удалось подобрать архетип для колоды: {picked_deck}")
                unknown_decks[tour].append(picked_deck)
                not_found_decklists.append(picked_deck)

        except Exception as e:
            print(f"Произошла ошибка: {e}")
            errors.append(e)

    print(f"Получилась следующая статистика пикнутых колод: {mapper}")
    print(f"Не получилось подобрать архетипы для следующих колод: {not_found_decklists}")
    print(f"Произошли следующие ошибки: {errors}")
    with pd.ExcelWriter(pathlib.Path("data", "total_picks.xlsx"), engine="xlsxwriter") as writer:
        for tour in mapper:
            lines = [[archetype, mapper[tour][archetype]] for archetype in mapper[tour]]
            lines.append(["Авторская сборка", len(unknown_decks[tour])])
            df = pd.DataFrame(sorted(lines, key=lambda x: x[1], reverse=True), columns=["Архетип", "Кол-во выбора"])
            df.to_excel(writer, sheet_name=f"{tour} тур", index=False)
            format_center = writer.book.add_format({'align': 'center'})
            worksheet = writer.sheets[f"{tour} тур"]
            worksheet.set_column('A:A', 60, format_center)
            worksheet.set_column('B:B', 15, format_center)

        total_mapper = defaultdict(int)
        for tour in mapper:
            for archetype in mapper[tour]:
                total_mapper[archetype] += mapper[tour][archetype]

        lines = [[archetype, total_mapper[archetype]] for archetype in total_mapper]
        df = pd.DataFrame(sorted(lines, key=lambda x: x[1], reverse=True), columns=["Архетип", "Кол-во выбора"])
        df.to_excel(writer, sheet_name="Общая статистика", index=False)
        format_center = writer.book.add_format({'align': 'center'})
        worksheet = writer.sheets["Общая статистика"]
        worksheet.set_column('A:A', 60, format_center)
        worksheet.set_column('B:B', 15, format_center)

        not_found_decklists_df = pd.DataFrame(not_found_decklists, columns=["Неизвестные архетипы"])
        not_found_decklists_df.to_excel(writer, sheet_name="Неизвестные архетипы", index=False)
        worksheet = writer.sheets["Неизвестные архетипы"]
        format_center = writer.book.add_format({'align': 'center'})
        worksheet.set_column('A:AAA', 80, format_center)


if __name__ == '__main__':
    analyze_turn()
