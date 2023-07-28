import requests
import pandas as pd


url = "https://admblag.kaiten.ru/api/latest/boards/390459/"
token = "d7167686-0d84-4bfb-8181-5fe4d1eef5f1"

# Set the Authorization header with the Bearer token
headers = {"Authorization": f"Bearer {token}"}

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

response = requests.get(url, headers=headers)

if response.status_code == 200:
    data = response.json()

    columns = {column["id"]: column["title"] for column in data.get("columns", [])}

    card_data_list = []
    cards = data.get("cards", [])
    for card in cards:
        card_id = card.get("id")
        card_url = f"https://admblag.kaiten.ru/api/latest/cards/{card_id}"
        card_response = requests.get(card_url, headers=headers)

        if card_response.status_code == 200:
            card_data = card_response.json()
            title = card_data.get("title")
            description = card_data.get("description")
            column_id = card_data.get("column_id")

            column_title = columns.get(column_id, "")

            comments = get_comments(card_id)

            card_data_list.append({
                "Card ID": card_id,
                "Title": title,
                "Description": description,
                "Column": column_title,
                "Comments": "".join([f"{comment[1]}: {comment[0]}\n" for comment in comments]),

            })

        else:
            print(f"Error: Failed to retrieve data for Card ID: {card_id}")

    # Convert the list of dictionaries to a DataFrame
    df = pd.DataFrame(card_data_list)

    # Export DataFrame to Excel with line wrapping enabled
    output_file = "output.xlsx"
    df.to_excel(output_file, index=False)  # Set wrap_text to True
    print(f"Data exported to {output_file}")
else:
    print(f"Error: Failed to retrieve data, status code: {response.status_code}")
