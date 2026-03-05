# 1. нужно установить python3, для него библиотеку selenium и для selenium chromedriver
# 2. заполнить константы EMAIL и PASSWORD для proberserk
# 3. в TOURNAMENT_LINK указать ссылку на турнир на proberserk
# 4. Если размещаешь по местам, то SORT_BY_PLACE = True, если по порядку сортировки, то SORT_BY_PLACE = False
# 5. Если не хочешь чтобы браузер открывался, укажи HEADLESS = True
# 6. создай в той же директории, где лежит upload_to_proberserk.py папку "players", в нее загружай деклисты,
#    название должно быть формата НОМЕРДЕКЛИСТА_ИМЯ_РЕЗУЛЬТАТ_ПОРЯДОКСОРТИРОВКИ/МЕСТО.txt.
#    деклист должен быть обычным файлом экспорта для proberserk

import os
import shutil
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils import login_to_proberserk, upload_deck_list, get_deck_name, create_deck, clean_deck_list, get_player_to_city_mapper, get_city_from_mapper, clean_player, get_deck_list


EMAIL = ""
PASSWORD = ""
TOURNAMENT_LINK = ""

DECKLISTS_FILEPATH = "deck_lists/"
UPLOADED_DECKLISTS_FILEPATH = DECKLISTS_FILEPATH + "uploaded/"
INVALID_DECKLISTS_FILEPATH = DECKLISTS_FILEPATH + "invalid/"

TIMEOUT = 60


def delete_deck_lists(only_duplicated: bool = False):

    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, timeout=TIMEOUT)

    login_to_proberserk(driver=driver, email=EMAIL, password=PASSWORD)

    response = requests.get(TOURNAMENT_LINK)
    response.raise_for_status()

    encountered_decks = []

    tournament_soup = BeautifulSoup(response.text, "html.parser")
    for deck in tournament_soup.find_all("tr"):
        cleaned_deck = deck.get_text().replace("\n", "").strip()
        if only_duplicated:
            if cleaned_deck not in encountered_decks:
                print(f'Вижу впервый раз колоду {cleaned_deck}, пропускаю ее')
                encountered_decks.append(cleaned_deck)
                continue
        print(f'Приступил к удалению колоды {cleaned_deck}')

        a_tag = deck.find("a")
        if not a_tag:
            continue

        driver.get(f"https://proberserk.ru{a_tag.get('href')}")

        forms = driver.find_elements(By.CSS_SELECTOR, 'form.d-inline[method="post"]')
        while forms:
            forms = driver.find_elements(By.CSS_SELECTOR, 'form.d-inline[method="post"]')
            old_html = driver.find_element(By.TAG_NAME, "html")
            form = forms[-1]

            btn = form.find_element(By.CSS_SELECTOR, 'button[name="delete-button"]')
            if "tournament" in driver.current_url:
                break

            btn.click()
            wait.until(EC.alert_is_present())
            Alert(driver).accept()

            wait.until(EC.staleness_of(form))
            wait.until(EC.staleness_of(old_html))

        print(f"Удалил колоду {cleaned_deck}")



def upload_deck_lists():

    driver = webdriver.Chrome()
    mapper = get_player_to_city_mapper()

    login_to_proberserk(driver=driver, email=EMAIL, password=PASSWORD)
    os.makedirs(UPLOADED_DECKLISTS_FILEPATH, exist_ok=True)
    os.makedirs(INVALID_DECKLISTS_FILEPATH, exist_ok=True)

    for path, folders, files in os.walk(DECKLISTS_FILEPATH):
        for deck_list_filename in files:
            if not deck_list_filename.endswith(".txt"):
                continue

            print(f"Приступил к обработке деклиста {deck_list_filename}")

            driver.get(TOURNAMENT_LINK)

            name, player, record, sort = deck_list_filename.split("_")
            sort = sort.replace(".txt", "")

            deck_list = get_deck_list(deck_list_filepath=(os.path.join(path, deck_list_filename)))
            if not deck_list:
                shutil.move(os.path.join(path, deck_list_filename), INVALID_DECKLISTS_FILEPATH)
                print(f"Деклист {deck_list} был пропущен из-за ошибки чтения!")
                continue

            name = get_deck_name(deck_list) or name
            player = clean_player(player=player)
            deck_list = clean_deck_list(deck_list)

            player_city = get_city_from_mapper(mapper=mapper, player=player)
            if player_city:
                player = f"{player} ({player_city})"

            create_deck(driver=driver, name=name, player=player, record=record, sort=sort)
            upload_deck_list(driver=driver, deck_list=deck_list)

            shutil.move(os.path.join(path, deck_list_filename), UPLOADED_DECKLISTS_FILEPATH)
            print(f"Перенес деклист {deck_list_filename} в папку uploaded")

            driver.get(TOURNAMENT_LINK)

    driver.quit()


if __name__ == '__main__':
    upload_deck_lists()
