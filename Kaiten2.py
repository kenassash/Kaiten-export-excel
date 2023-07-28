import json
import pandas as pd
import requests

token = "d7167686-0d84-4bfb-8181-5fe4d1eef5f1"
headers = {"Authorization": f"Bearer {token}"}


def get_board_to_space(headers):
    url = "https://admblag.kaiten.ru/api/latest/spaces/157246/boards"
    board_response = requests.get(url, headers=headers)
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


# Call the function to retrieve the data
board_ids, board_titles, column_ids, column_titles = get_board_to_space(headers)


def get_card(board_id, column_id, headers):
    url = f"https://admblag.kaiten.ru/api/latest/boards/{board_id}/"
    card_response = requests.get(url, headers=headers)
    if card_response.status_code == 200:
        card_data = card_response.json()
        cards = card_data.get("cards", [])
        card_ids = [card.get("id") for card in cards if card.get("column_id") == column_id]
        return card_ids


def get_card_description(card_ids, headers):
    titles = []
    descriptions = []
    for card_id in card_ids:
        card_url = f"https://admblag.kaiten.ru/api/latest/cards/{card_id}"
        card_response = requests.get(card_url, headers=headers)
        if card_response.status_code == 200:
            card_data = card_response.json()
            title = card_data.get("title")
            description = card_data.get("description")
            titles.append(title)
            descriptions.append(description)
    return titles, descriptions


def get_comments(card_id):
    comments_url = f"https://admblag.kaiten.ru/api/latest/cards/{card_id}/comments"
    comments_response = requests.get(comments_url, headers=headers)

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


# Initialize empty lists to store data
all_board_ids = []
all_board_titles = []
all_column_ids = []
all_column_titles = []
all_card_ids = []
all_titles = []
all_descriptions = []
all_comments = []

# Iterate over board_ids and column_ids together using zip()
for board_id, column_id in zip(board_ids, column_ids):
    card_ids = get_card(board_id, column_id, headers)
    if card_ids is not None:  # Check if card_ids is not None before appending
        all_board_ids.extend([board_id] * len(card_ids))
        all_board_titles.extend([board_titles[board_ids.index(board_id)]] * len(card_ids))
        all_column_ids.extend([column_id] * len(card_ids))
        all_column_titles.extend([column_titles[column_ids.index(column_id)]] * len(card_ids))
        all_card_ids.extend(card_ids)

        titles, descriptions = get_card_description(card_ids, headers)
        all_titles.extend(titles)
        all_descriptions.extend(descriptions)

        # Retrieve and store comments for each card
        for card_id in card_ids:
            comments = get_comments(card_id)
            all_comments.append("".join([f"{comment[1]}: {comment[0]}\n" for comment in comments]))

# Create a DataFrame from the lists
data = {
    "Board Id": all_board_ids,
    "Board Title": all_board_titles,
    "Column Id": all_column_ids,
    "Column Title": all_column_titles,
    "Cards Id": all_card_ids,
    "Title": all_titles,
    "Description": all_descriptions,
    "Comments": all_comments,  # Add the Comments list to the data dictionary
}
df = pd.DataFrame(data)

# Use the explode function to split the card_ids list into separate rows
df = df.explode("Cards Id", ignore_index=True)

# Write DataFrame to Excel file
df.to_excel("output.xlsx", index=False)
