import argparse

import pandas as pd

from pathlib import Path


DEFAULT_START_CASE_ID = 141202168
DEFAULT_END_CASE_ID = 141202300

OUTPUT_FILE = Path(
    "data/input/case_ids.xlsx"
)

PROGRESS_FILE = Path(
    "progress.txt"
)

RESULTS_FILE = Path(
    "data/output/cases_results.xlsx"
)

RETRY_FILE = Path(
    "data/output/retry_queue.xlsx"
)

DONE_FILE = Path(
    "data/output/DONE.txt"
)


def reset_state():

    files = [
        PROGRESS_FILE,
        RESULTS_FILE,
        RETRY_FILE,
        DONE_FILE,
    ]

    for file in files:

        if file.exists():

            file.unlink()

            print(
                f"Deleted: {file}"
            )

    PROGRESS_FILE.write_text(
        "0"
    )

    print(
        "Progress reset to 0"
    )


def generate_case_ids(
    start_case_id,
    end_case_id,
):

    if end_case_id < start_case_id:

        raise ValueError(
            "END_CASE_ID must be "
            "greater than or equal "
            "to START_CASE_ID."
        )

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    case_ids = list(
        range(
            start_case_id,
            end_case_id + 1,
        )
    )

    df = pd.DataFrame(
        {
            "case_id": case_ids
        }
    )

    df.to_excel(
        OUTPUT_FILE,
        index=False,
    )

    print(
        f"Generated "
        f"{len(case_ids):,} Case IDs"
    )

    print(
        f"Start: {start_case_id}"
    )

    print(
        f"End:   {end_case_id}"
    )

    print(
        f"Saved to: {OUTPUT_FILE}"
    )


def main():

    parser = argparse.ArgumentParser(
        description=(
            "Generate Philadelphia "
            "Civil Docket Case IDs."
        )
    )

    parser.add_argument(
        "start",
        nargs="?",
        type=int,
        default=DEFAULT_START_CASE_ID,
        help="First Case ID",
    )

    parser.add_argument(
        "end",
        nargs="?",
        type=int,
        default=DEFAULT_END_CASE_ID,
        help="Last Case ID",
    )

    parser.add_argument(
        "--reset-state",
        action="store_true",
        help=(
            "Delete previous progress, "
            "results and retry queue."
        ),
    )

    args = parser.parse_args()

    if args.reset_state:

        reset_state()

    generate_case_ids(
        args.start,
        args.end,
    )


if __name__ == "__main__":
    main()