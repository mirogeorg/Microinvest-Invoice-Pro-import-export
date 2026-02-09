try:
    from .manager import ExcelSQLManager
except ImportError:
    from manager import ExcelSQLManager


def main():
    app = ExcelSQLManager()
    app.run()


if __name__ == '__main__':
    main()
