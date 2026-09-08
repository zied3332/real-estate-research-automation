import subprocess
import time
import random
import sys

from pathlib import Path


MAIN_SCRIPT = Path("src/main.py")

DONE_FILE = Path(
    "data/output/DONE.txt"
)

PROGRESS_FILE = Path(
    "progress.txt"
)

RETRY_FILE = Path(
    "data/output/retry_queue.xlsx"
)


# Wait between complete browser sessions.
MIN_WAIT_BETWEEN_RUNS = 25
MAX_WAIT_BETWEEN_RUNS = 50

# If main.py crashes, wait longer.
CRASH_WAIT_MIN = 60
CRASH_WAIT_MAX = 120


def run_main():

    print()
    print("=" * 60)
    print(
        "Starting next scraper session..."
    )
    print("=" * 60)
    print()

    result = subprocess.run(
        [
            sys.executable,
            str(MAIN_SCRIPT),
        ]
    )

    print()
    print("=" * 60)
    print(
        f"main.py finished "
        f"with code: "
        f"{result.returncode}"
    )
    print("=" * 60)
    print()

    return result.returncode


def show_progress():

    if PROGRESS_FILE.exists():

        progress = (
            PROGRESS_FILE
            .read_text()
            .strip()
        )

        print(
            f"Current progress index: "
            f"{progress}"
        )


def main():

    run_number = 1

    while True:

        if DONE_FILE.exists():

            print()
            print("=" * 60)
            print(
                "ALL CASES FINISHED"
            )
            print("=" * 60)

            break

        print()
        print(
            f"Runner session #{run_number}"
        )

        show_progress()

        code = run_main()

        if DONE_FILE.exists():

            print()
            print(
                "Scraper reports that "
                "all cases are finished."
            )

            break

        if code != 0:

            wait_time = random.randint(
                CRASH_WAIT_MIN,
                CRASH_WAIT_MAX,
            )

            print(
                f"main.py crashed."
            )

            print(
                f"Waiting {wait_time}s "
                f"before restarting..."
            )

            time.sleep(
                wait_time
            )

            run_number += 1

            continue

        wait_time = random.randint(
            MIN_WAIT_BETWEEN_RUNS,
            MAX_WAIT_BETWEEN_RUNS,
        )

        print(
            f"Session complete."
        )

        print(
            f"Waiting {wait_time}s "
            f"before next session..."
        )

        time.sleep(
            wait_time
        )

        run_number += 1


if __name__ == "__main__":
    main()