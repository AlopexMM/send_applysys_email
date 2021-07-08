import gi

gi.require_version("Gtk", "3.0")
from gi.repository import Gtk
from .applysys import process_excel

class mainWindow(Gtk.Window):
    def __init__(self):
        super().__init__(title="Scanntech")
        self.xls_path = ""

        box = Gtk.Box(spacing=6)
        self.add(box)

        choose_file_button = Gtk.Button(label="Seleccionar Planilla")
        choose_file_button.connect('clicked', self.on_file_clicked)
        box.add(choose_file_button)

        send_mail_button = Gtk.Button(label="Enviar emails")
        send_mail_button.connect('clicked', self.on_send_mail_clicked)
        box.add(send_mail_button)

        self.label_mail_sended = Gtk.Label()
        self.label_mail_sended.set_text("")
        box.add(self.label_mail_sended)

    def on_file_clicked(self, widget):
        dialog = Gtk.FileChooserDialog(
            title="Seleccionar la planilla", parent=self, action=Gtk.FileChooserAction.OPEN
        )
        dialog.add_buttons(
            Gtk.STOCK_CANCEL,
            Gtk.ResponseType.CANCEL,
            Gtk.STOCK_OPEN,
            Gtk.ResponseType.OK,
        )

        response = dialog.run()

        if response == Gtk.ResponseType.OK:
            self.xls_path = dialog.get_filename()

        dialog.destroy()

    def on_send_mail_clicked(self, widget):
        process_excel(self.xls_path)
        self.label_mail_sended.set_text("Enviados")

def main():
    win = mainWindow()
    win.connect("destroy", Gtk.main_quit)
    win.show_all()
    Gtk.main()

if __name__ == "__main__":
    main()
