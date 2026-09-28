"""
Step 2 - Load the CSV files into a SQLite database and run SQL queries.

SQLite is a file-based database that ships with Python, so nothing needs
to be installed. The same SQL works with small changes on PostgreSQL,
SQL Server or Snowflake.
"""
import re
import sqlite3

import pandas as pd

import config

TABLES = ["suppliers", "products", "weekly_demand", "purchase_orders", "inventory_snapshot"]
SQL_FILE = config.ROOT / "sql" / "analysis_queries.sql"


def build_database() -> None:
    """Create (or overwrite) data/supply_chain.db from the CSV files."""
    with sqlite3.connect(config.DB_PATH) as conn:
        for table in TABLES:
            df = pd.read_csv(config.DATA_DIR / f"{table}.csv")
            df.to_sql(table, conn, if_exists="replace", index=False)
    print(f"Loaded {len(TABLES)} tables into {config.DB_PATH.name}")


def load_queries() -> dict[str, str]:
    """Split the .sql file into {query_name: sql_text}."""
    text = SQL_FILE.read_text(encoding="utf-8")
    parts = re.split(r"^-- name:\s*(\w+)\s*$", text, flags=re.MULTILINE)
    # parts = [header, name1, sql1, name2, sql2, ...]
    return {parts[i]: parts[i + 1].strip() for i in range(1, len(parts), 2)}


def run_query(name: str) -> pd.DataFrame:
    sql = load_queries()[name]
    with sqlite3.connect(config.DB_PATH) as conn:
        return pd.read_sql_query(sql, conn)


if __name__ == "__main__":
    build_database()
    print(run_query("supplier_delivery_performance").to_string(index=False))
