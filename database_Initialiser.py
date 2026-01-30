from dotenv import load_dotenv
from utils import get_conn

load_dotenv()

def initialize_all_databases():
    conn = get_conn()
    cursor = conn.cursor()

    initialize_active_contracts_database(cursor)
    initialize_contracts_database(cursor)
    initialize_closed_contracts_database(cursor)
    initialize_given_and_additional_database(cursor)
    initialize_paid_principle_and_paid_percentage_database(cursor)
    initialize_paid_principle_registry_database(cursor)
    initialize_outflow_order_database(cursor)
    initialize_outflow_in_registry_database(cursor)
    initialize_adding_percent_amount_database(cursor)
    initialize_paid_percent_amount_database(cursor)
    initialize_inflow_order_only_principal_amount_database(cursor)
    initialize_inflow_order_only_percent_amount_database(cursor)
    initialize_blk_list_database(cursor)
    initialize_inflow_order_both_database(cursor)

    conn.commit()
    conn.close()



# Initializing databases for money control window tables
def initialize_contracts_database(cursor):

        cursor.execute("""
                      CREATE TABLE IF NOT EXISTS contracts (
                          unique_id SERIAL PRIMARY KEY,
                          contract_id INTEGER,
                          contract_open_date TEXT,
                          first_percent_payment_date TEXT,
                          name_surname TEXT,
                          id_number TEXT,
                          tel_number TEXT,
                          item_name TEXT,
                          model TEXT,
                          IMEI TEXT,
                          trusted_person TEXT,
                          comment TEXT,
                          given_money INTEGER,
                          percent_day_quantity INTEGER,
                          first_added_percent NUMERIC,
                          sum_of_principle_and_percent NUMERIC GENERATED ALWAYS AS (given_money + first_added_percent) STORED,
                          office_mob_number TEXT
                      )
                  """)

        cursor.execute("DROP VIEW IF EXISTS contracts_view")

        cursor.execute("""
            CREATE VIEW contracts_view AS
            SELECT
                unique_id,
                contract_id,
                contract_open_date,
                first_percent_payment_date,
                name_surname,
                id_number,
                tel_number,
                item_name,
                model,
                IMEI,
                trusted_person,
                comment,
                given_money,
                percent_day_quantity,
                first_added_percent,
                (given_money + first_added_percent) AS sum_of_principle_and_percent,
                office_mob_number
            FROM contracts
        """)



def initialize_closed_contracts_database(cursor):

        cursor.execute("""
                      CREATE TABLE IF NOT EXISTS closed_contracts (
                          id INTEGER PRIMARY KEY,
                          contract_open_date TEXT,
                          name_surname TEXT,
                          id_number TEXT,
                          tel_number TEXT,
                          item_name TEXT,
                          model TEXT,
                          IMEI TEXT,
                          trusted_person TEXT,
                          comment TEXT,
                          percent NUMERIC,
                          percent_day_quantity INTEGER,
                          given_money INTEGER,
                          additional_money INTEGER,
                          paid_principle NUMERIC,
                          added_percents NUMERIC,
                          paid_percents NUMERIC,
                          status TEXT,
                          date_of_closing TEXT
                      )
                  """)



def initialize_active_contracts_database(cursor):

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS active_contracts (
                id SERIAL PRIMARY KEY,
                date TEXT,
                days_after_C_O INTEGER NOT NULL DEFAULT 0,
                name_surname TEXT,
                id_number TEXT,
                tel_number TEXT,
                item_name TEXT,
                model TEXT,
                imei TEXT,
                type TEXT,
                trusted_person TEXT,
                comment TEXT,
                given_money NUMERIC NOT NULL DEFAULT 0,
                percent NUMERIC NOT NULL DEFAULT 0,
                day_quantity INTEGER NOT NULL DEFAULT 0,
                additional_amounts NUMERIC NOT NULL DEFAULT 0,
                principal_paid NUMERIC NOT NULL DEFAULT 0,
                principal_should_be_paid NUMERIC GENERATED ALWAYS AS (
                    given_money + additional_amounts - principal_paid
                ) STORED,
                added_percents NUMERIC NOT NULL DEFAULT 0,
                paid_percents NUMERIC NOT NULL DEFAULT 0,
                percent_should_be_paid NUMERIC GENERATED ALWAYS AS (
                    added_percents - paid_percents
                ) STORED,
                is_visible TEXT DEFAULT 'აქტიური'
            )
        """)

        # Create or replace the view that exposes generated columns explicitly
        cursor.execute("DROP VIEW IF EXISTS active_contracts_view")
        cursor.execute("""
               CREATE VIEW active_contracts_view AS
               SELECT
                   id,
                   date,
                   days_after_C_O,
                   name_surname,
                   id_number,
                   tel_number,
                   item_name,
                   model,
                   imei,
                   type,
                   trusted_person,
                   comment,
                   given_money,
                   percent,
                   day_quantity,
                   additional_amounts,
                   principal_paid,
                   (given_money + additional_amounts - principal_paid) AS principal_should_be_paid,
                   added_percents,
                   paid_percents,
                   (added_percents - paid_percents) AS percent_should_be_paid,
                   is_visible
               FROM active_contracts
           """)

def initialize_given_and_additional_database(cursor):

        cursor.execute("""
                      CREATE TABLE IF NOT EXISTS given_and_additional_database (
                          unique_id SERIAL PRIMARY KEY,
                          contract_id INTEGER,
                          date_of_outflow TEXT,
                          name_surname TEXT,
                          amount NUMERIC,
                          status TEXT
                      )
                  """)

def initialize_paid_principle_and_paid_percentage_database(cursor):

        cursor.execute("""
                      CREATE TABLE IF NOT EXISTS paid_principle_and_paid_percentage_database (
                          unique_id SERIAL PRIMARY KEY,
                          contract_id INTEGER,
                          date_of_inflow TEXT,
                          name_surname TEXT,
                          amount NUMERIC,
                          status TEXT
                      )
                  """)


def initialize_paid_principle_registry_database(cursor):

        cursor.execute("""
                      CREATE TABLE IF NOT EXISTS paid_principle_registry (
                          unique_id SERIAL PRIMARY KEY,
                          contract_id INTEGER,
                          date_of_C_O TEXT,
                          name_surname TEXT,
                          tel_number TEXT,
                          id_number TEXT,
                          item_name TEXT,
                          model TEXT,
                          IMEI TEXT,
                          given_money INTEGER,
                          date_of_payment TEXT,
                          payment_amount NUMERIC,
                          status TEXT
                      )
                  """)



def initialize_outflow_order_database(cursor):

        cursor.execute("""
                      CREATE TABLE IF NOT EXISTS outflow_order (
                          unique_id SERIAL PRIMARY KEY,
                          contract_id INTEGER,
                          name_surname TEXT,
                          tel_number TEXT,
                          amount NUMERIC,
                          date TEXT,
                          status TEXT
                      )
                  """)



def initialize_outflow_in_registry_database(cursor):

        cursor.execute("""
                      CREATE TABLE IF NOT EXISTS outflow_in_registry (
                          unique_id SERIAL PRIMARY KEY,
                          contract_id INTEGER,
                          date_of_C_O TEXT,
                          name_surname TEXT,
                          tel_number TEXT,
                          id_number TEXT,
                          item_name TEXT,
                          model TEXT,
                          IMEI TEXT,
                          given_money INTEGER,
                          date_of_addition TEXT,
                          additional_amount INTEGER,
                          status TEXT
                      )
                  """)


def initialize_adding_percent_amount_database(cursor):

        cursor.execute("""
                              CREATE TABLE IF NOT EXISTS adding_percent_amount (
                                  unique_id SERIAL PRIMARY KEY,
                                  contract_id INTEGER,
                                  date_of_C_O TEXT,
                                  name_surname TEXT,
                                  tel_number TEXT,
                                  id_number TEXT,
                                  item_name TEXT,
                                  model TEXT,
                                  IMEI TEXT,
                                  date_of_percent_addition TEXT,
                                  percent_amount INTEGER,
                                  status TEXT
                              )
                          """)


def initialize_paid_percent_amount_database(cursor):

        cursor.execute(""" CREATE TABLE IF NOT EXISTS paid_percent_amount (
                                  unique_id SERIAL PRIMARY KEY,
                                  contract_id INTEGER,
                                  date_of_C_O TEXT,
                                  name_surname TEXT,
                                  tel_number TEXT,
                                  id_number TEXT,
                                  item_name TEXT,
                                  model TEXT,
                                  IMEI TEXT,
                                  set_date TEXT,
                                  date_of_percent_addition TEXT,
                                  paid_amount INTEGER,
                                  status TEXT
                              )
                          """)



def initialize_inflow_order_only_principal_amount_database(cursor):

        cursor.execute("""
                              CREATE TABLE IF NOT EXISTS inflow_order_only_principal_amount (
                                  unique_id SERIAL PRIMARY KEY,
                                  contract_id INTEGER,
                                  name_surname TEXT,
                                  principle_paid_amount NUMERIC,
                                  payment_date TEXT,
                                  sum_of_money_paid NUMERIC
                              )
                          """)



def initialize_inflow_order_only_percent_amount_database(cursor):

        cursor.execute("""
                              CREATE TABLE IF NOT EXISTS inflow_order_only_percent_amount (
                                  unique_id SERIAL PRIMARY KEY,
                                  contract_id INTEGER,
                                  name_surname TEXT,
                                  payment_date TEXT,
                                  set_date TEXT,
                                  percent_paid_amount INTEGER,
                                  sum_of_money_paid NUMERIC
                              )
                          """)


def initialize_blk_list_database(cursor):

        cursor.execute("""
                   CREATE TABLE IF NOT EXISTS black_list (
                       id SERIAL PRIMARY KEY,
                       name_surname TEXT,
                       id_number TEXT,
                       tel_number TEXT,
                       imei TEXT
                   )
               """)

def initialize_inflow_order_both_database(cursor):

        cursor.execute("""
                CREATE TABLE IF NOT EXISTS inflow_order_both (
                    unique_id SERIAL PRIMARY KEY,
                    contract_id INTEGER,
                    name_surname TEXT,
                    payment_date TEXT,
                    principle_paid_amount NUMERIC NOT NULL DEFAULT 0,
                    percent_paid_amount NUMERIC NOT NULL DEFAULT 0,
                    sum_of_money_paid NUMERIC GENERATED ALWAYS AS (
                        principle_paid_amount + percent_paid_amount
                    ) STORED
                )
            """)

        # Create a view for the table
        cursor.execute("DROP VIEW IF EXISTS inflow_order_both_view")

        cursor.execute("""
                CREATE VIEW inflow_order_both_view AS
                SELECT
                    unique_id,
                    contract_id,
                    name_surname,
                    payment_date,
                    principle_paid_amount,
                    percent_paid_amount,
                    sum_of_money_paid
                FROM inflow_order_both
            """)