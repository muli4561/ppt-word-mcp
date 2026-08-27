"""Run the HTTP service with ``python -m ppt_word_gen``."""

import uvicorn


def main() -> None:
    uvicorn.run("ppt_word_gen.app:app", host="0.0.0.0", port=8000)


if __name__ == "__main__":
    main()
