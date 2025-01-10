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
