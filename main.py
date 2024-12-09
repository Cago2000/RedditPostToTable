import praw
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font


# Function to create a Reddit instance
def create_reddit_instance(client_id, client_secret, user_agent):
    reddit = praw.Reddit(
        client_id=client_id,
        client_secret=client_secret,
        user_agent=user_agent
    )
    return reddit


def get_submissions(reddit, subreddit, limit, sort_type):
    match sort_type:
        case 'hot':
            submissions = reddit.subreddit(subreddit).hot(limit=limit+1)
        case'new':
            submissions = reddit.subreddit(subreddit).new(limit=limit+1)
        case 'rising':
            submissions = reddit.subreddit(subreddit).rising(limit=limit+1)
        case 'top':
            submissions = reddit.subreddit(subreddit).top(limit=limit+1)
        case _:
            submissions = None
    return submissions


def load_submission_data_to_table(submissions, keywords):
    rows = []
    post_number = 1
    for submission in submissions:
        post_has_keyword = False
        if submission.author == 'PokeUpdateBot':
            continue
        for keyword in keywords:
            if keyword in submission.title.lower() or keyword in submission.selftext.lower():
                post_has_keyword = True
        rows.append({
            "Username": str(submission.author),
            "Content": f'Post {post_number}|-- {submission.title}:\n{submission.selftext.replace('\n', ' ')}',
            "URL": submission.url,
            "Depth": 0
        })
        post_number += 1

        submission.comments.replace_more(limit=0)  # Load all comments

        # Process top-level comments
        comment_number = 1  # Reset comment number for each post
        conversation_has_keyword = False
        for top_level_comment in submission.comments:
            current_rows = extract_comments(submission, top_level_comment, keywords, conversation_has_keyword, depth=2, comment_number=comment_number)
            comment_number += 1
        if len(current_rows) > 0: 
            rows.extend(current_rows)

            rows.append({
            "Username": "",
            "Content": "",
            "Depth": 0
            })
            rows.append({
                "Username": "",
                "Content": "",
                "Depth": 0
            })

        else:
            if not post_has_keyword:
                rows.pop(-1)
    return rows


def extract_comments(submission, comment, keywords, conversation_has_keyword, depth=1, comment_number=1):
    comments_data = []
    indent_type = "Comment" if depth > 1 else "Post"
    indent = f'{indent_type} {comment_number} |-- '
    indented_comment = '    ' * (depth - 1) + indent + comment.body.replace('\n', ' ')  # Remove unnecessary line breaks
    for keyword in keywords:
        if keyword.lower() in comment.body.lower() or conversation_has_keyword:
            conversation_has_keyword = True
            comments_data.append({
                "Username": str(comment.author),
                "Content": indented_comment,
                "URL": f"https://www.reddit.com{submission.permalink}{comment.id}",
                "Depth": depth
        })
    reply_number = 1
    for reply in comment.replies:
        current_row = extract_comments(submission, reply, keywords, conversation_has_keyword, depth + 1, reply_number) # Recurse into replies
        comments_data.extend(current_row)
        reply_number += 1
    return comments_data


# Function to search Reddit posts and save them with comments in a tree structure to an Excel file
def save_data_to_xlsx(reddit, subreddit, keywords, limit, sort_type, filename="reddit_posts_tree.xlsx"):
    submissions = get_submissions(reddit, subreddit, limit, sort_type)
    rows = load_submission_data_to_table(submissions, keywords)
    df = pd.DataFrame(rows)
    wb = Workbook()
    ws = wb.active
    ws.append(["Username", "Content"])
    colors = ['F0F0F0', 'D3D3D3', 'B0C4DE',  'ADD8E6',  'E6E6FA',  'FFFACD',  'F5DEB3',  'FAFAD2',  'E0FFFF',  'F5F5F5']
    for _, row in df.iterrows():
        depth = row["Depth"]
        color_index = depth % len(colors)
        if depth == 0:
            ws.append([row["Username"], row["Content"], row["URL"]])
        else:
            fill = PatternFill(start_color=colors[color_index], end_color=colors[color_index], fill_type="solid")
            ws.append([row["Username"], row["Content"], row["URL"]])
            ws.cell(row=ws.max_row, column=2).fill = fill
            ws.cell(row=ws.max_row, column=2).alignment = Alignment(wrap_text=True)  # Ensure text is wrapped nicely

    # Auto-adjust column widths
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    make_keyword_cells_bold_in_cells(keywords, ws)
    wb.save(filename)
    print(f"Data saved to {filename}")


def make_keyword_cells_bold_in_cells(keywords, ws):
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                # Check each keyword
                for keyword in keywords:
                    if keyword.lower() in cell.value.lower():
                        # Change the entire cell to bold
                        cell.font = Font(bold=True)
                        break  # Exit after the first keyword is found


def get_client_credentials():
    with open('client_credentials.txt', 'r') as file:
        for line in file:
            name, value = line.strip().split('=')
            if name == 'client_id':
                client_id = value
            elif name == 'client_secret':
                client_secret = value
        return client_id, client_secret


def main():
    # Reddit API credentials
    client_id, client_secret = get_client_credentials()
    user_agent = "reddit search by crusader"
    reddit = create_reddit_instance(client_id, client_secret, user_agent)

    subreddit = 'pokemon'
    keywords = ["pokemon"]
    sort_type = 'hot'  # Options: 'hot', 'new', 'rising', 'top'
    limit = 30  # Amount of submissions being considered

    # Search posts and save them to a CSV
    save_data_to_xlsx(reddit, subreddit, keywords, limit=limit, sort_type=sort_type, filename="reddit_posts.xlsx")


if __name__ == "__main__":
    main()

