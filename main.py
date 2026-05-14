import time
from halifax import halifax_runner


def main() -> None:
    halifax_runner()


if __name__ == "__main__":
    start_time = time.perf_counter()
    main()
    elapsed_time = time.perf_counter() - start_time
    print(f"Execution time: {elapsed_time:.2f} seconds")
