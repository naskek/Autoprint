"""Entry point for launching the BarTender GUI application."""
from bt_app.gui import App


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
