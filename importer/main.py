try:
    from .manager import run_app
except ImportError:
    from manager import run_app


def main():
    run_app()


if __name__ == '__main__':
    main()
