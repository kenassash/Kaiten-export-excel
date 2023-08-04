# -*- coding: utf-8 -*-
import json
import pandas as pd
import requests


class KaitenDataProcessor:
    def __init__(self, token):
        self.token = token
        self.headers = {"Authorization": f"Bearer {self.token}"}
        self.data = {
            "Board Id": [],
            "Board Title": [],
            "Column Id": [],
            "Column Title": [],
            "Cards Id": [],
            "Title": [],
            "Description": [],
            "Comments": [],
        }

    def get_board_to_space(self):
        print("Выполняется функция get_board_to_space()")
        url = "https://admblag.kaiten.ru/api/latest/spaces/157246/boards"
        board_response = requests.get(url, headers=self.headers)
        if board_response.status_code == 200:
            data = board_response.json()
            # Create lists to store data
            board_ids = []
            board_titles = []
            column_ids = []
            column_titles = []

            # Iterate through each board in the list
            for src in data:
                board_id = src["id"]
                board_title = src["title"]

                columns = src.get("columns", [])
                for column in columns:
                    column_id = column["id"]
                    column_title = column["title"]

                    # Append data to the lists
                    board_ids.append(board_id)
                    board_titles.append(board_title)
                    column_ids.append(column_id)
                    column_titles.append(column_title)
            return board_ids, board_titles, column_ids, column_titles

    def get_card(self, board_id, column_id):
        print(f"Выполняется функция get_card() для board_id={board_id}, column_id={column_id}")
        url = f"https://admblag.kaiten.ru/api/latest/boards/{board_id}/"
        card_response = requests.get(url, headers=self.headers)
        if card_response.status_code == 200:
            card_data = card_response.json()
            cards = card_data.get("cards", [])
            card_ids = [card.get("id") for card in cards if card.get("column_id") == column_id]
            return card_ids

    def get_card_description(self, card_ids):
        print(f"Выполняется функция get_card_description() для card_ids={card_ids}")
        titles = []
        descriptions = []
        for card_id in card_ids:
            card_url = f"https://admblag.kaiten.ru/api/latest/cards/{card_id}"
            card_response = requests.get(card_url, headers=self.headers)
            if card_response.status_code == 200:
                card_data = card_response.json()
                title = card_data.get("title")
                description = card_data.get("description")
                titles.append(title)
                descriptions.append(description)
        return titles, descriptions

    def get_comments(self, card_id):
        print(f"Выполняется функция get_comments() для card_id={card_id}")
        comments_url = f"https://admblag.kaiten.ru/api/latest/cards/{card_id}/comments"
        comments_response = requests.get(comments_url, headers=self.headers)

        if comments_response.status_code == 200:
            comments_data = comments_response.json()
            comments_list = []
            for comment in comments_data:
                text = comment.get("text")
                full_name = comment.get("author", {}).get("full_name")
                comments_list.append((text, full_name))
            return comments_list
        else:
            print(f"Error: Failed to retrieve comments for Card ID: {card_id}")
            return []

    def process_data(self):
        print("Начинается обработка данных")
        # Call the function to retrieve the data
        board_ids, board_titles, column_ids, column_titles = self.get_board_to_space()

        for board_id, column_id in zip(board_ids, column_ids):
            print(f"Обрабатывается board_id={board_id}, column_id={column_id}")
            card_ids = self.get_card(board_id, column_id)
            if card_ids is not None:
                self.data["Board Id"].extend([board_id] * len(card_ids))
                self.data["Board Title"].extend([board_titles[board_ids.index(board_id)]] * len(card_ids))
                self.data["Column Id"].extend([column_id] * len(card_ids))
                self.data["Column Title"].extend([column_titles[column_ids.index(column_id)]] * len(card_ids))
                self.data["Cards Id"].extend(card_ids)

                titles, descriptions = self.get_card_description(card_ids)
                self.data["Title"].extend(titles)
                self.data["Description"].extend(descriptions)

                for card_id in card_ids:
                    comments = self.get_comments(card_id)
                    self.data["Comments"].append("".join([f"{comment[1]}: {comment[0]}\n" for comment in comments]))

        print("Завершение обработки данных")

    def to_dataframe(self):
        print("Создание DataFrame")
        df = pd.DataFrame(self.data)
        return df

    def to_excel(self, file_path):
        print(f"Сохранение в Excel: {file_path}")
        df = self.to_dataframe()
        df.to_excel(file_path, index=False)


# Usage example:
if __name__ == "__main__":
    token = "d7167686-0d84-4bfb-8181-5fe4d1eef5f1"
    data_processor = KaitenDataProcessor(token)
    data_processor.process_data()
    data_processor.to_excel("output.xlsx")
    input("Нажмите Enter для завершения...")
