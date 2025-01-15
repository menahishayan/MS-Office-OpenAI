"""
This project attempts to create an open source implementation of OpenAI and Gemini to 
work with MS Office 365 Apps such as Excel, Word, and Powerpoint.

For questions, comments, or concerns, please raise a PR or Issue on the GitHub repository.

Licensed under GNU General Public License v3.0. See LICENSE for more information.
"""

import os
import sys

from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()


def main():
    # Get arguments from VBA
    if len(sys.argv) != 2:
        print("Error: Missing arguments. Usage: script.py <prompt>")
        sys.exit(1)

    prompt = sys.argv[1]

    # Call OpenAI API
    try:
        client = OpenAI(
            api_key=os.environ.get(
                "OPENAI_API_KEY"
            ),  # This is the default and can be omitted
        )

        chat_completion = client.chat.completions.create(
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            model="gpt-4o",
        )
        print(chat_completion.choices[0].message.content)
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
