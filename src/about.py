# -*- coding: utf-8 -*-

from .libs import dt, Qt, QVBoxLayout, QHBoxLayout, QLabel, QDialog, QFrame, open_new


class About(QDialog):
    def __init__(self, version, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle("About TkAstroDb")
        self.box = QVBoxLayout(self)
        self.setLayout(self.box)
        self.box.setAlignment(Qt.AlignHCenter)
        self.header = QLabel(self)
        self.header.setText("TkAstroDb")
        self.header.setStyleSheet("font: 25px Monospace; border: 1px solid #343a40; border-radius: 5px; padding: 10px;")
        self.header.setAlignment(Qt.AlignCenter)
        self.box.addWidget(self.header)
        self.create_info(version)

    def create_info(self, version):
        info = {
            "Version": version,
            "Date Built": "21.12.2018",
            "Date Updated": dt.now().strftime("%d.%m.%Y"),
            "Thanks To": "Alois Trendl, Flavia Alonzo, Sjoerd Visser",
            "Developed By": "Tanberk Celalettin Kutlu",
            "Contact": "tckutlu@gmail.com",
            "GitHub": "https://github.com/dildeolupbiten"
        }
        for key, value in info.items():
            frame = QFrame(self)
            frame.box = QHBoxLayout(frame)
            frame.setLayout(frame.box)
            title = QLabel(frame)
            title.setText(key)
            title.setFixedWidth(150)
            title.setStyleSheet("border: 1px solid #343a40; border-radius: 5px; padding: 10px; background-color: #343a40")
            description = QLabel(frame)
            description.setText(value)
            name = key.replace(" ", "_")
            description.setObjectName(name)
            description.setStyleSheet(
                """
                #wid {
                    color: white; 
                    margin-left: 5px;
                }

                #wid:hover {
                    border: 1px solid #343a40; 
                    border-radius: 5px; 
                    padding: 10px;
                    color: white;
                }
                """.replace("wid", name)
            )
            if key in ["Contact", "GitHub"]:
                description.mouseReleaseEvent = lambda e, v=value: self.open_url(v)
                description.setCursor(Qt.PointingHandCursor)
            frame.box.addWidget(title)
            frame.box.addWidget(description)
            self.box.addWidget(frame)

    @staticmethod
    def open_url(url):
        open_new(url)
