

import oracledb

USERNAME = "testuser"

PASSWORD = "your_password"

# Oracle TNS 접속 문자열

TNS = """

(DESCRIPTION=

    (ADDRESS=(PROTOCOL=TCP)(HOST=192.0.2.10)(PORT=1521))

    (CONNECT_DATA=(SERVICE_NAME=ORCLPDB1))

)

"""

try:

    print("[INFO] Oracle DB 연결 시도...")

    with oracledb.connect(

        user=USERNAME,

        password=PASSWORD,

        dsn=TNS

    ) as connection:

        print("[SUCCESS] Oracle DB 연결 성공!")

        with connection.cursor() as cursor:

            cursor.execute("""

                SELECT owner, table_name

                FROM all_tables

                WHERE owner NOT IN ('SYS', 'SYSTEM')

                ORDER BY owner, table_name

            """)

            tables = cursor.fetchall()

            print(f"[INFO] 조회 가능한 테이블 목록: {len(tables)}개")

            for owner, table_name in tables:

                print(f"{owner}.{table_name}")

except oracledb.Error as e:

    print(f"[ERROR] Oracle DB 오류: {e}")