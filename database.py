#!/usr/bin/env python
# coding: utf-8
import sqlite3


# =============================================================================
#  SqlDB - handle Database for SPC Master
# =============================================================================
class SqlDB():
    # SQLite database file name
    dbname: str = None

    # Transaction flag
    OK = None
    ERRORMSG = None

    def __init__(self, dbname):
        self.dbname = dbname

    # -------------------------------------------------------------------------
    #  put
    #  execute SQL
    #
    #  argument:
    #    sql : SQL statement
    # -------------------------------------------------------------------------
    def put(self, sql):
        con = sqlite3.connect(self.dbname)
        cur = con.cursor()

        try:
            cur.execute(sql)
            con.commit()
            self.OK = True
            self.ERRORMSG = None
        except Exception as e:
            print(e)
            self.OK = False
            self.ERRORMSG = e

        con.close()

    # -------------------------------------------------------------------------
    #  get
    #  query with SQL
    #
    #  argument:
    #    sql : SQL statement
    #
    #  return
    #    out : matrix of output
    # -------------------------------------------------------------------------
    def get(self, sql):
        con = sqlite3.connect(self.dbname)
        cur = con.cursor()

        try:
            cur.execute(sql)
            out = cur.fetchall()
            self.OK = True
            self.ERRORMSG = None
        except Exception as e:
            print(e)
            self.OK = False
            self.ERRORMSG = e

        con.close()
        return out

    # -------------------------------------------------------------------------
    #  sql
    #  create sql replacing ?s by parameters
    #
    #  argument:
    #    sentense   : SQL statement with ?s
    #    parameters : parameters to replace
    #
    #  return
    #    sentense : full SQL replaced with parameters
    # -------------------------------------------------------------------------
    def sql(self, sentense, parameters):
        for param in parameters:
            sentense = sentense.replace('?', str(param), 1)
        return sentense
