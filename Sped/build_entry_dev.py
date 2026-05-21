from __future__ import annotations

import os

from Revisor import main


if __name__ == "__main__":
    os.environ["SPED_ENV"] = "dev"
    main()
