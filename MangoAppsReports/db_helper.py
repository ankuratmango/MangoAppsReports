import mysql.connector
import pandas as pd
from typing import List, Tuple, Any, Generator, Optional

class DatabaseConnection:
    def __init__(self, host: str, user: str, password: str, database: str):
        self.host = host
        self.user = user
        self.password = password
        self.database = database
        self.connection: Optional[mysql.connector.MySQLConnection] = None

    def connect(self):
        if self.connection is None:
            self.connection = mysql.connector.connect(
                host=self.host,
                user=self.user,
                password=self.password,
                database=self.database
            )

    def disconnect(self):
        if self.connection:
            self.connection.close()
            self.connection = None

    def execute_query(self, query: str, params: Tuple[Any, ...] = ()) -> None:
        if not self.connection:
            raise ValueError("Database is not connected.")

        cursor = self.connection.cursor()
        cursor.execute(query, params)
        self.connection.commit()
        cursor.close()

    def fetch_one(self, query: str, params: Tuple[Any, ...] = ()) -> Optional[Tuple[Any, ...]]:
        if not self.connection:
            raise ValueError("Database is not connected.")

        cursor = self.connection.cursor()
        cursor.execute(query, params)
        result = cursor.fetchone()
        cursor.close()
        return result

    def fetch_all(self, query: str, params: Tuple[Any, ...] = ()) -> List[Tuple[Any, ...]]:
        if not self.connection:
            raise ValueError("Database is not connected.")

        cursor = self.connection.cursor()
        cursor.execute(query, params)
        results = cursor.fetchall()
        column_names = [desc[0] for desc in cursor.description]
        cursor.close()
        return pd.DataFrame(results, columns=column_names)

    def fetch_in_batches(self, query: str, params: Tuple[Any, ...] = (), batch_size: int = 100) -> Generator[List[Tuple[Any, ...]], None, None]:
        if not self.connection:
            raise ValueError("Database is not connected.")

        cursor = self.connection.cursor()
        cursor.execute(query, params)

        while True:
            batch = cursor.fetchmany(batch_size)
            if not batch:
                break
            yield batch

        cursor.close()

    def delete_rows(self, query: str, params: Tuple[Any, ...] = ()) -> None:
        self.execute_query(query, params)

