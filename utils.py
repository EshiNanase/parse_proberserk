import json
import re
from pprint import pprint
import requests
import pathlib
from collections import defaultdict

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

TIMEOUT = 90

re_file = re.compile(r"drive\.google\.com/file/d/([^/]+)/")
re_uc = re.compile(r"[?&]id=([^&]+)")
re_open = re.compile(r"drive\.google\.com/open\?id=([^&]+)")

NXT_DATA_FILEPATH = "data/nxt_data.json"


def clean_player(player: str) -> str:
    return player.strip().replace("ё", "е").strip()


def get_player_to_city_mapper():
    df = pd.read_excel('data/players.xlsx', sheet_name='Список участников')
    mapper = {}
    for row in df.values:
        mapper[clean_player(row[0])] = row[-1]
    print("Открыл df с участниками турнира и сделал мапер")
    pprint(mapper)
    return mapper


def get_city_from_mapper(mapper: dict, player: str):
    city = mapper.get(clean_player(player))
    if city:
        print(f"Нашел для {player} город {city}")
        return city
    print(f"Не нашел город для {player}")


def create_wait(driver: webdriver.Chrome):
    return WebDriverWait(driver=driver, timeout=TIMEOUT)


def get_deck_name(deck_list: list[str]) -> str | None:
    if "#" in deck_list[0]:
        name = deck_list[0].replace("\n", "").replace("#", "")
        name_arr = list(name)
        name_arr[0] = name_arr[0].upper()
        return "".join(name_arr)


def get_deck_list(deck_list_filepath: str) -> list[str]:

    try:

        with open(deck_list_filepath, "r", encoding="utf-8") as deck_file:
            deck_list = deck_file.read()

        deck_list_tts = json.loads(deck_list)
        deck_list_map = defaultdict(int)
        for card in deck_list_tts["ObjectStates"][0]["ContainedObjects"]:
            deck_list_map[clean_player(card["Nickname"])] += 1
        deck_list_arr = [f"{quantity} {card}" for card, quantity in deck_list_map.items()]
        print("Составил деклист из формата TTS")

    except json.decoder.JSONDecodeError:
        deck_list_arr = [clean_player(card) for card in deck_list.split("\n")]

    except Exception as e:
        print(f"Ошибка при загрузке деклиста {deck_list_filepath}: {e}")
        return []

    print(f"Деклист был успешно получен: {', '.join(deck_list_arr)}")
    return deck_list_arr


def clean_deck_list(deck_list: list[str]) -> str:
    if "#" in deck_list[0]:
        deck_list.pop(0)
    return "\n".join(deck_list)


def create_deck(driver: webdriver.Chrome, name: str, player: str, record: str, sort: str):
    wait = create_wait(driver=driver)
    create_deck_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href^="/deck/new/"]')))
    create_deck_link.click()

    generate_hash_button = driver.find_element(By.ID, "generate-hash")
    generate_hash_button.click()

    name_input = wait.until(EC.presence_of_element_located((By.NAME, 'deck[name]')))
    name_input.send_keys(name)

    player_input = wait.until(EC.presence_of_element_located((By.NAME, 'deck[player]')))
    player_input.send_keys(player)

    record_input = wait.until(EC.presence_of_element_located((By.NAME, 'deck[record]')))
    record_input.send_keys(record)

    place_input = wait.until(EC.presence_of_element_located((By.NAME, "deck[place]")))
    place_input.send_keys(sort)

    sort_input = wait.until(EC.presence_of_element_located((By.NAME, "deck[sort]")))
    sort_input.send_keys(sort)

    save_form_button = driver.find_element(By.NAME, "save-button")
    save_form_button.click()

    print(f"Создал колоду {name} для игрока {player} с результатом {record} и местом {sort}")


def upload_deck_list(driver: webdriver.Chrome, deck_list: str):
    wait = create_wait(driver=driver)
    add_deck_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href^="/deck/add/card_list/"]')))
    add_deck_button.click()

    cards_input = wait.until(EC.presence_of_element_located((By.ID, 'deck_card_list_list')))
    cards_input.send_keys(deck_list)

    save_form_button = driver.find_element(By.NAME, "save-button")
    save_form_button.click()

    deck_list_arr = deck_list.split("\n")

    print(f"Добавил следующие карты: {', '.join(deck_list_arr)}")


def login_to_proberserk(driver: webdriver.Chrome, email: str, password: str):
    wait = create_wait(driver=driver)
    driver.get('https://proberserk.ru/login')

    email_input = wait.until(EC.presence_of_element_located((By.NAME, 'email')))
    password_input = wait.until(EC.presence_of_element_located((By.NAME, 'password')))

    email_input.send_keys(email)
    password_input.send_keys(password)

    submit_button = driver.find_element(By.NAME, "submit_button")
    submit_button.click()

    print(f"Вошел в аккаунт {email}")


def extract_id(url: str) -> str | None:
    for rx in (re_file, re_uc, re_open):
        m = rx.search(url)
        if m:
            return m.group(1)
    return

def download_public(file_id: str, dst_path: pathlib.Path, chunk=32768):
    url = "https://drive.usercontent.google.com/download"
    params = {"export": "download", "id": file_id, "confirm": "t"}
    with requests.Session() as s, s.get(url, params=params, stream=True) as response:
        response.raise_for_status()
        with open(dst_path, "wb") as f:
            for chunk_bytes in response.iter_content(chunk_size=chunk):
                if chunk_bytes:
                    f.write(chunk_bytes)


def download_form_data():
    form_data_filepath = "data/form_data.xlsx"
    spreadsheet = "1shhE_d3Y6TVRjfW_tr9WQvkJ5aQ7IYPw1Sps1RZsbmM"
    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet}/export?format=xlsx"
    resp = requests.get(url, timeout=60)
    resp.raise_for_status()
    with open(form_data_filepath, "wb") as f:
        f.write(resp.content)
    print(f"Скачал ответы по форме с {url}")


def get_mapper_from_nxt_data():
    with open(NXT_DATA_FILEPATH, "r", encoding="utf-8") as file:
        data = json.load(file)
    mapper = {}
    for line in data:
        name = clean_player(line.pop("name"))
        mapper[name] = {key: value for key, value in line.items()}
    print(f"Создал маппер из данных nxt для {len(mapper)} карт")
    return mapper
