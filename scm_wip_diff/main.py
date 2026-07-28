import tkinter as tk

from scm_wip_diff.gui import App


def main():
    root = tk.Tk()
    root.title("GTK WIP 일일 비교")
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
