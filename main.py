import sys

from single_instance import acquire_single_instance_lock


if __name__ == "__main__":
    lock = acquire_single_instance_lock()
    if lock is None:
        sys.exit(0)

    from gui_app import main

    main(lock)
