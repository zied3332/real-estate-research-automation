import subprocess
import time
import sys
from pathlib import Path

MAIN_SCRIPT = Path("src/main.py")
WAIT_BETWEEN_RUNS = 10

def run_main():
    print("\n==============================")
    print("Starting main.py for next 10 cases...")
    print("==============================\n")

    result = subprocess.run([sys.executable, str(MAIN_SCRIPT)])

    print("\n==============================")
    print(f"main.py finished with code: {result.returncode}")
    print("==============================\n")

    return result.returncode


def main():
    while True:
        code = run_main()

        if code != 0:
            print("main.py crashed or stopped with error.")
            print("Stopping runner.py so you can check the problem.")
            break

        print(f"Waiting {WAIT_BETWEEN_RUNS} seconds before next batch...")
        time.sleep(WAIT_BETWEEN_RUNS)


if __name__ == "__main__":
    main()