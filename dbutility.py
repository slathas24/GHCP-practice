"""
sqlite_utility.py

A generic SQLite CLI utility for running SQL commands against any database
(default: holders.db) - not tied to a specific table or schema.

Commands:
    tables                          List all tables in the database
    schema TABLE                    Show the CREATE TABLE statement for a table
    schema                          Show CREATE statements for all tables
    query "SELECT ..."              Run a read-only query, print results as a table
    exec "INSERT/UPDATE/DELETE ..." Run a write statement (INSERT/UPDATE/DELETE/etc.)
    exec "DROP/ALTER/..." --confirm Destructive DDL statements require --confirm
    shell                           Interactive prompt: type SQL, get results, repeat

Safety notes:
    - 'query' only accepts statements that don't modify data (enforced by
      running them in a read-only connection, so even a disguised write
      inside a query fails rather than silently succeeding).
    - 'exec' is for INSERT/UPDATE/DELETE. DROP/ALTER/TRUNCATE-style statements
      additionally require --confirm before running.
    - Every write runs inside a transaction that's rolled back automatically
      if anything raises partway through.
    - This tool executes whatever SQL string you give it - it does not
      parameterize values for you. Don't build query/exec strings from
      untrusted external input (e.g. raw user form submissions) without
      your own parameterization/allow-listing on top of this.
"""

import sqlite3
import argparse
import sys
import os

DB_PATH_DEFAULT = "holders.db"

# Statement types considered destructive enough to require --confirm
DESTRUCTIVE_KEYWORDS = ("DROP", "ALTER", "TRUNCATE", "VACUUM")


# ----------------------------
# Connections
# ----------------------------

def get_write_connection(db_path: str) -> sqlite3.Connection:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


def get_readonly_connection(db_path: str) -> sqlite3.Connection:
    """
    Opens the DB in SQLite's true read-only mode (mode=ro) via URI, so any
    write attempted through this connection fails at the SQLite level -
    not just by convention.
    """
    if not os.path.exists(db_path):
        raise FileNotFoundError(f"Database file not found: {db_path}")
    uri = f"file:{db_path}?mode=ro"
    conn = sqlite3.connect(uri, uri=True)
    conn.row_factory = sqlite3.Row
    return conn


# ----------------------------
# tables / schema
# ----------------------------

def list_tables(conn: sqlite3.Connection):
    rows = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name;"
    ).fetchall()
    if not rows:
        print("No tables found.")
        return
    print("Tables:")
    for r in rows:
        print(f"  - {r['name']}")


def show_schema(conn: sqlite3.Connection, table: str = None):
    if table:
        row = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name = ?", (table,)
        ).fetchone()
        if row is None:
            print(f"Error: table '{table}' not found.")
            sys.exit(1)
        print(row["sql"])
    else:
        rows = conn.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name;"
        ).fetchall()
        if not rows:
            print("No tables found.")
            return
        for r in rows:
            print(r["sql"] + ";\n")


# ----------------------------
# Generic query (read-only)
# ----------------------------

def print_rows(rows):
    if not rows:
        print("(no rows)")
        return

    headers = rows[0].keys()
    widths = {h: max(len(h), max(len(str(r[h])) for r in rows)) for h in headers}

    print("  ".join(h.ljust(widths[h]) for h in headers))
    print("  ".join("-" * widths[h] for h in headers))
    for r in rows:
        print("  ".join(str(r[h]).ljust(widths[h]) for h in headers))

    print(f"\n{len(rows)} row(s).")


def run_query(db_path: str, sql: str):
    try:
        conn = get_readonly_connection(db_path)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        sys.exit(1)

    try:
        cursor = conn.execute(sql)
        rows = cursor.fetchall()
        print_rows(rows)
    except sqlite3.Error as e:
        print(f"Query error: {e}")
        sys.exit(1)
    finally:
        conn.close()


# ----------------------------
# Generic exec (write)
# ----------------------------

def is_destructive(sql: str) -> bool:
    first_word = sql.strip().split(None, 1)[0].upper() if sql.strip() else ""
    return first_word in DESTRUCTIVE_KEYWORDS


def run_exec(db_path: str, sql: str, confirm: bool):
    if is_destructive(sql) and not confirm:
        print(f"Refusing to run destructive statement without --confirm:\n  {sql}")
        sys.exit(1)

    conn = get_write_connection(db_path)
    try:
        cursor = conn.execute(sql)
        conn.commit()
        print(f"OK. {cursor.rowcount if cursor.rowcount != -1 else 0} row(s) affected.")
    except sqlite3.Error as e:
        conn.rollback()
        print(f"Execution error (rolled back): {e}")
        sys.exit(1)
    finally:
        conn.close()


# ----------------------------
# Interactive shell
# ----------------------------

def run_shell(db_path: str):
    print(f"SQLite shell on '{db_path}'. Type SQL statements, or 'exit' to quit.")
    print("Prefix a destructive statement (DROP/ALTER/TRUNCATE/VACUUM) is allowed here without --confirm,")
    print("since you're typing it interactively yourself - be careful.\n")

    conn = get_write_connection(db_path)
    try:
        while True:
            try:
                sql = input("sql> ").strip()
            except (EOFError, KeyboardInterrupt):
                print()
                break

            if not sql:
                continue
            if sql.lower() in ("exit", "quit"):
                break

            try:
                cursor = conn.execute(sql)
                if sql.strip().upper().startswith("SELECT") or sql.strip().upper().startswith("PRAGMA"):
                    rows = cursor.fetchall()
                    print_rows(rows)
                else:
                    conn.commit()
                    print(f"OK. {cursor.rowcount if cursor.rowcount != -1 else 0} row(s) affected.")
            except sqlite3.Error as e:
                conn.rollback()
                print(f"Error: {e}")
    finally:
        conn.close()
        print("Shell closed.")


# ----------------------------
# CLI
# ----------------------------

def main():
    parser = argparse.ArgumentParser(description="Generic SQLite command utility")
    parser.add_argument("--db", default=DB_PATH_DEFAULT,
                         help=f"Path to the SQLite database (default: {DB_PATH_DEFAULT})")
    subparsers = parser.add_subparsers(dest="command", required=True)

    subparsers.add_parser("tables", help="List all tables")

    schema_parser = subparsers.add_parser("schema", help="Show CREATE TABLE statement(s)")
    schema_parser.add_argument("table", nargs="?", default=None,
                                help="Table name (omit to show all tables)")

    query_parser = subparsers.add_parser("query", help="Run a read-only SQL query (e.g. SELECT)")
    query_parser.add_argument("sql", help="SQL statement to run")

    exec_parser = subparsers.add_parser("exec", help="Run a write SQL statement (INSERT/UPDATE/DELETE/...)")
    exec_parser.add_argument("sql", help="SQL statement to run")
    exec_parser.add_argument("--confirm", action="store_true",
                              help="Required for destructive statements (DROP/ALTER/TRUNCATE/VACUUM)")

    subparsers.add_parser("shell", help="Interactive SQL prompt")

    args = parser.parse_args()

    if args.command == "tables":
        conn = get_write_connection(args.db)
        try:
            list_tables(conn)
        finally:
            conn.close()

    elif args.command == "schema":
        conn = get_write_connection(args.db)
        try:
            show_schema(conn, args.table)
        finally:
            conn.close()

    elif args.command == "query":
        run_query(args.db, args.sql)

    elif args.command == "exec":
        run_exec(args.db, args.sql, args.confirm)

    elif args.command == "shell":
        run_shell(args.db)


if __name__ == "__main__":
    main()