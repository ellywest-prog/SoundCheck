# -*- coding: utf-8 -*-
import sys
import os

devnull = open(os.devnull, 'w')
old_stderr = sys.stderr
sys.stderr = devnull

from PyQt6.QtWidgets import QApplication
app = QApplication(sys.argv)

sys.stderr = old_stderr
devnull.close()

import asyncio
import numpy as np
from PyQt6.QtWidgets import (QMainWindow, QVBoxLayout,
                             QWidget, QLabel, QHBoxLayout, QPushButton,
                             QMessageBox, QSlider, QSizePolicy)
from PyQt6.QtCore import QTimer, Qt, QThread, pyqtSignal, QSize, QPointF, QRectF, QSettings
from PyQt6.QtGui import (QFont, QColor, QKeySequence, QShortcut,
                         QPainter, QPixmap, QBrush, QPen, QPainterPath, QIcon, QPolygonF)
import pyqtgraph as pg

import pyaudiowpatch as pa


# ─────────────────────────────────────────────────────────────────────
# Custom monochrome icon renderer (QPainter, no emoji / color font)
# ─────────────────────────────────────────────────────────────────────
def _mk_icon(shape: str, sz: int = 34) -> QIcon:
    """Draw a single-color gun-metal media icon onto a transparent pixmap."""
    px = QPixmap(sz, sz)
    px.fill(Qt.GlobalColor.transparent)
    p = QPainter(px)
    p.setRenderHint(QPainter.RenderHint.Antialiasing)

    c  = QColor(168, 182, 195)   # silver-grey
    p.setBrush(QBrush(c))
    p.setPen(Qt.PenStyle.NoPen)

    m   = sz * 0.13              # outer margin
    inn = sz - 2 * m             # inner area size
    bw  = sz * 0.10              # vertical-bar width
    cy  = sz / 2.0               # vertical centre

    def tri_r(x0, y0, x1, y1):
        """Filled triangle pointing RIGHT."""
        q = QPolygonF([QPointF(x0, y0), QPointF(x0, y1), QPointF(x1, (y0+y1)/2)])
        p.drawPolygon(q)

    def tri_l(x0, y0, x1, y1):
        """Filled triangle pointing LEFT."""
        q = QPolygonF([QPointF(x1, y0), QPointF(x1, y1), QPointF(x0, (y0+y1)/2)])
        p.drawPolygon(q)

    def vbar(x, w=None):
        w = w or bw
        p.drawRoundedRect(QRectF(x, m, w, inn), 1.5, 1.5)

    def speaker_body():
        """Draw a small speaker cone; returns x where waves start."""
        # Rectangle base
        rb_w, rb_h = inn * 0.18, inn * 0.45
        rb_x = m + inn * 0.04
        rb_y = cy - rb_h / 2
        p.drawRoundedRect(QRectF(rb_x, rb_y, rb_w, rb_h), 1, 1)
        # Cone (trapezoid pointing right)
        cone_x0 = rb_x + rb_w
        cone_x1 = rb_x + rb_w + inn * 0.30
        poly = QPolygonF([
            QPointF(cone_x0, cy - rb_h * 0.50),
            QPointF(cone_x0, cy + rb_h * 0.50),
            QPointF(cone_x1, cy + rb_h * 0.85),
            QPointF(cone_x1, cy - rb_h * 0.85),
        ])
        p.drawPolygon(poly)
        return cone_x1 + sz * 0.05   # wave start x

    # ── Shape definitions ─────────────────────────────────────────────
    if shape == 'prev':           # |◄
        vbar(m)
        tri_l(m + bw + sz*0.06, m, sz - m, sz - m)

    elif shape == 'play_pause':   # ►‖
        tw = inn * 0.44
        tri_r(m, m, m + tw, sz - m)
        gap = sz * 0.07
        bx  = m + tw + gap
        bw2 = inn * 0.13
        vbar(bx, bw2)
        vbar(bx + bw2 + gap, bw2)

    elif shape == 'stop':         # ■
        r = inn * 0.10
        p.drawRoundedRect(QRectF(m, m, inn, inn), r, r)

    elif shape == 'next':         # ►|
        tri_r(m, m, sz - m - bw - sz*0.06, sz - m)
        vbar(sz - m - bw)

    elif shape == 'mute':
        wx = speaker_body()
        # X mark
        pen = QPen(c, sz * 0.085, Qt.PenStyle.SolidLine,
                   Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        xr = sz * 0.13
        p.drawLine(QPointF(wx,      cy - xr), QPointF(wx + 2*xr, cy + xr))
        p.drawLine(QPointF(wx+2*xr, cy - xr), QPointF(wx,        cy + xr))

    elif shape in ('vol_down', 'vol_up'):
        wx = speaker_body()
        waves = 1 if shape == 'vol_down' else 2
        pen = QPen(c, sz * 0.07, Qt.PenStyle.SolidLine,
                   Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
        p.setPen(pen)
        p.setBrush(Qt.BrushStyle.NoBrush)
        for i in range(waves):
            r  = sz * (0.14 + i * 0.15)
            rc = QRectF(wx - r + sz*0.01, cy - r, r * 2, r * 2)
            p.drawArc(rc, -50 * 16, 100 * 16)

    p.end()
    return QIcon(px)

# ── Windows Media Session ─────────────────────────────────────────────
try:
    from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager
    WINSDK_AVAILABLE = True
except ImportError:
    WINSDK_AVAILABLE = False

# ── System Volume Control (pycaw) ─────────────────────────────────────
try:
    from pycaw.pycaw import AudioUtilities
    PYCAW_AVAILABLE = True

    # Cache the AudioDevice object — EndpointVolume is a direct property
    _spk_device = None

    def _ensure_vol():
        global _spk_device
        if _spk_device is None:
            try:
                _spk_device = AudioUtilities.GetSpeakers()
            except Exception:
                _spk_device = None
        return _spk_device

    def get_master_volume():
        try:
            d = _ensure_vol()
            return int(d.EndpointVolume.GetMasterVolumeLevelScalar() * 100) if d else 50
        except Exception:
            return 50

    def set_master_volume(pct):
        try:
            d = _ensure_vol()
            if d:
                d.EndpointVolume.SetMasterVolumeLevelScalar(
                    max(0.0, min(1.0, pct / 100.0)), None
                )
        except Exception:
            pass

except ImportError:
    PYCAW_AVAILABLE = False
    def get_master_volume(): return 50
    def set_master_volume(pct): pass


# ── ctypes media / volume keys ────────────────────────────────────────
try:
    import ctypes
    CTYPES_AVAILABLE = True
except ImportError:
    CTYPES_AVAILABLE = False

VK_MEDIA_NEXT_TRACK = 0xB0
VK_MEDIA_PREV_TRACK = 0xB1
VK_MEDIA_STOP       = 0xB2
VK_MEDIA_PLAY_PAUSE = 0xB3
VK_VOLUME_MUTE      = 0xAD
VK_VOLUME_DOWN      = 0xAE
VK_VOLUME_UP        = 0xAF

def send_media_key(vk_code):
    if CTYPES_AVAILABLE:
        ctypes.windll.user32.keybd_event(vk_code, 0, 0x0000, 0)
        ctypes.windll.user32.keybd_event(vk_code, 0, 0x0002, 0)


# ─────────────────────────────────────────────────────────────────────
class MediaInfoThread(QThread):
    info_ready = pyqtSignal(str, str)

    def run(self):
        if not WINSDK_AVAILABLE:
            return
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        while not self.isInterruptionRequested():
            try:
                title, artist = loop.run_until_complete(self.fetch())
                self.info_ready.emit(title, artist)
            except Exception:
                pass
            self.msleep(2000)

    async def fetch(self):
        manager = await GlobalSystemMediaTransportControlsSessionManager.request_async()
        session = manager.get_current_session()
        if session:
            info   = await session.try_get_media_properties_async()
            title  = info.title  if info.title  else "Bilinmeyen Şarkı"
            artist = info.artist if info.artist else "Bilinmeyen Sanatçı"
            return (title, artist)
        return ("Müzik Çalmıyor", "-")


# ─────────────────────────────────────────────────────────────────────
class RealTimeSpekApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real-Time Audio Quality Analyzer")
        self.resize(870, 530)

        self.p = None
        self.stream = None
        self.fs = 48000
        self.fft_size = 8192
        self.audio_data = np.zeros(self.fft_size)
        self.audio_lock = False

        self.current_title  = "Bekleniyor..."
        self.current_artist = "Sanatçı bilgisi yok"

        self.num_bands        = 32
        self.band_edges       = None
        self.band_centers     = None
        self.smoothed_bands   = np.zeros(self.num_bands)
        self.peaks            = np.zeros(self.num_bands)
        self.peak_hold_frames = np.zeros(self.num_bands)
        self.peak_hold_max    = 30
        self.peak_drop_rate   = 1.0
        self.smoothed_db      = 0.0
        self.db_peak          = 0.0
        self.db_peak_hold_frames = 0
        self.bar_items_cyan   = []
        self.bar_items_yellow = []
        self.bar_items_red    = []
        self.peak_items       = []
        self.led_step         = 2.0

        # Fixed brushes — set once
        self.BAR_BRUSH  = pg.mkBrush(QColor(0, 210, 220))
        self.PEAK_BRUSH = pg.mkBrush(QColor(255, 255, 255))
        self.DB_PEAK_GREEN  = pg.mkBrush(QColor(50, 230, 130))
        self.DB_PEAK_YELLOW = pg.mkBrush(QColor(255, 175, 40))
        self.DB_PEAK_RED    = pg.mkBrush(QColor(255, 55, 55))

        # Volume slider sync rate limiter
        self._vol_sync_counter = 0

        self.init_ui()
        self.setup_audio()
        self.setup_shortcuts()
        
        self.settings = QSettings("SoundcheckApp", "BandSettings")
        saved_bands = self.settings.value("band_count", 16, type=int)
        self.set_bands(saved_bands)

        # Sync slider to actual system volume once at startup
        if PYCAW_AVAILABLE:
            self.vol_slider.setValue(get_master_volume())

        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_ui)
        self.update_timer.start(20)   # 50 fps

        if WINSDK_AVAILABLE:
            self.media_thread = MediaInfoThread()
            self.media_thread.info_ready.connect(self.update_media_labels)
            self.media_thread.start()

    # ─────────────────────────────────────────────────────────────────
    def init_ui(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #0D0F17; }

            QLabel#title {
                color: #FFFFFF; font-size: 22px; font-weight: bold;
                font-family: 'Segoe UI', sans-serif;
            }
            QLabel#artist {
                color: #7A82A0; font-size: 13px;
                font-family: 'Segoe UI', sans-serif;
            }
            QLabel#status {
                color: #3E4552; font-size: 10px;
                font-family: 'Segoe UI', sans-serif;
            }
            QLabel#vol_label {
                color: #6A7890; font-size: 10px;
                font-family: 'Segoe UI', sans-serif;
                min-width: 28px;
            }
            QWidget#top_bar {
                background-color: #13161F;
                border-bottom: 1px solid #1E2130; border-radius: 6px;
            }
            QWidget#ctrl_bar {
                background-color: #13161F;
                border-radius: 6px; border: 1px solid #1E2130;
            }

            /* ── Tüm butonlarda focus halkasını kapat ── */
            QPushButton:focus {
                outline: none;
                border: 1px solid #2A2F40;
            }

            /* ── Band seçim butonları ── */
            QPushButton {
                background-color: #1E2130; color: #8A91A6;
                padding: 5px 11px; border-radius: 4px;
                font-weight: bold; font-size: 11px;
                border: 1px solid #2A2F40;
            }
            QPushButton:hover   { background-color: #262C42; color: #D0D8F0; }
            QPushButton:pressed { background-color: #161A28; }

            /* ── Medya butonları — Gun-Metal 3D ── */
            QPushButton#media_btn {
                background: qlineargradient(
                    x1:0, y1:0, x2:0, y2:1,
                    stop:0.00 #5A6470,
                    stop:0.12 #4A5560,
                    stop:0.88 #2C353C,
                    stop:1.00 #1A2025
                );
                color: #B8C4CC;
                font-size: 22px;
                border-top:    1px solid #7A8A96;
                border-left:   1px solid #606C78;
                border-right:  1px solid #404C54;
                border-bottom: 3px solid #0E1518;
                border-radius: 7px;
                min-width: 46px;
                min-height: 38px;
                padding: 0px 4px;
            }
            QPushButton#media_btn:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:0, y2:1,
                    stop:0.00 #68767E,
                    stop:0.15 #566470,
                    stop:0.85 #36424A,
                    stop:1.00 #222C34
                );
                color: #D8E4EC;
                border-top:    1px solid #8A9AA6;
                border-left:   1px solid #707C88;
                border-right:  1px solid #505C64;
                border-bottom: 3px solid #0A1215;
            }
            QPushButton#media_btn:pressed {
                background: qlineargradient(
                    x1:0, y1:0, x2:0, y2:1,
                    stop:0.00 #141A1E,
                    stop:0.20 #1E2830,
                    stop:1.00 #2A3840
                );
                color: #90A0AA;
                border-top:    3px solid #0A1215;
                border-left:   1px solid #303C44;
                border-right:  1px solid #404C54;
                border-bottom: 1px solid #5A6870;
                padding-top: 4px;
            }
            /* Focus halkası media butonlarda da kapalı */
            QPushButton#media_btn:focus {
                outline: none;
                border-top:    1px solid #7A8A96;
                border-left:   1px solid #606C78;
                border-right:  1px solid #404C54;
                border-bottom: 3px solid #0E1518;
            }

            /* ── Volume slider ── */
            QSlider#vol_slider {
                min-height: 20px;
            }
            QSlider#vol_slider::groove:horizontal {
                height: 5px;
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #1A2028, stop:1 #252E38);
                border: 1px solid #303C48;
                border-radius: 3px;
            }
            QSlider#vol_slider::sub-page:horizontal {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #1A6080, stop:1 #00B8D4);
                border-radius: 3px;
                height: 5px;
            }
            QSlider#vol_slider::add-page:horizontal {
                background: #1A2028;
                border: 1px solid #283038;
                border-radius: 3px;
                height: 5px;
            }
            QSlider#vol_slider::handle:horizontal {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #7A9AAA, stop:1 #4A6878);
                border: 1px solid #8AAABB;
                border-bottom: 2px solid #283038;
                width: 14px; height: 14px;
                margin: -5px 0;
                border-radius: 7px;
            }
            QSlider#vol_slider::handle:horizontal:hover {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #9ABACA, stop:1 #5A8898);
                border: 1px solid #A0CCDD;
            }
            QSlider#vol_slider::handle:horizontal:pressed {
                background: qlineargradient(x1:0,y1:0,x2:0,y2:1,
                    stop:0 #3A5868, stop:1 #6A8898);
            }
        """)

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(12, 12, 12, 8)
        layout.setSpacing(8)

        # ── Top bar ──────────────────────────────────────────────────
        top_bar = QWidget(); top_bar.setObjectName("top_bar")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(12, 8, 12, 8)

        placeholder = QWidget(); placeholder.setFixedSize(30, 30)
        top_layout.addWidget(placeholder)

        center_w = QWidget()
        center_l = QVBoxLayout(center_w)
        center_l.setContentsMargins(0, 0, 0, 0); center_l.setSpacing(2)

        self.lbl_title = QLabel(self.current_title)
        self.lbl_title.setObjectName("title")
        self.lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.lbl_artist = QLabel(self.current_artist)
        self.lbl_artist.setObjectName("artist")
        self.lbl_artist.setAlignment(Qt.AlignmentFlag.AlignCenter)

        center_l.addWidget(self.lbl_title)
        center_l.addWidget(self.lbl_artist)
        top_layout.addWidget(center_w, stretch=1)

        self.btn_info = QPushButton("?")
        self.btn_info.setFixedSize(30, 30)
        self.btn_info.clicked.connect(self.show_info)
        top_layout.addWidget(self.btn_info)
        layout.addWidget(top_bar)

        # ── Control bar ───────────────────────────────────────────────
        ctrl_bar = QWidget(); ctrl_bar.setObjectName("ctrl_bar")
        ctrl_layout = QHBoxLayout(ctrl_bar)
        ctrl_layout.setContentsMargins(10, 5, 10, 5)
        ctrl_layout.setSpacing(5)

        # ── Transport buttons (Prev / Play-Pause / Stop / Next) ───────
        transport_defs = [
            ('prev',       "Önceki Parça  [\u2190]",   lambda: send_media_key(VK_MEDIA_PREV_TRACK)),
            ('play_pause', "Oynat / Dur  [Space]",  lambda: send_media_key(VK_MEDIA_PLAY_PAUSE)),
            ('stop',       "Durdur  [S]",           lambda: send_media_key(VK_MEDIA_STOP)),
            ('next',       "Sonraki Parça  [\u2192]",  lambda: send_media_key(VK_MEDIA_NEXT_TRACK)),
        ]
        self.media_btn_refs = []
        for shape, tip, action in transport_defs:
            btn = QPushButton()
            btn.setIcon(_mk_icon(shape, 34))
            btn.setIconSize(QSize(34, 34))
            btn.setObjectName("media_btn")
            btn.setToolTip(tip)
            btn.setFixedSize(52, 44)
            btn.clicked.connect(action)
            ctrl_layout.addWidget(btn)
            self.media_btn_refs.append(btn)

        # ── Separator ─────────────────────────────────────────────────
        sep = QLabel("│")
        sep.setStyleSheet("color:#2A3040; font-size:20px; margin: 0 4px;")
        ctrl_layout.addWidget(sep)

        # ── Mute button ───────────────────────────────────────────────
        btn_mute = QPushButton()
        btn_mute.setIcon(_mk_icon('mute', 34))
        btn_mute.setIconSize(QSize(34, 34))
        btn_mute.setObjectName("media_btn")
        btn_mute.setToolTip("Sessiz / Aç  [M]")
        btn_mute.setFixedSize(52, 44)
        btn_mute.clicked.connect(lambda: send_media_key(VK_VOLUME_MUTE))
        ctrl_layout.addWidget(btn_mute)
        self.media_btn_refs.append(btn_mute)

        # ── Volume Down ───────────────────────────────────────────────
        btn_vol_dn = QPushButton()
        btn_vol_dn.setIcon(_mk_icon('vol_down', 34))
        btn_vol_dn.setIconSize(QSize(34, 34))
        btn_vol_dn.setObjectName("media_btn")
        btn_vol_dn.setToolTip("Ses Azalt  [\u2212]")
        btn_vol_dn.setFixedSize(52, 44)
        btn_vol_dn.clicked.connect(self._vol_down)
        ctrl_layout.addWidget(btn_vol_dn)
        self.media_btn_refs.append(btn_vol_dn)

        # ── Volume Slider ─────────────────────────────────────────────
        self.vol_slider = QSlider(Qt.Orientation.Horizontal)
        self.vol_slider.setObjectName("vol_slider")
        self.vol_slider.setRange(0, 100)
        self.vol_slider.setValue(50)
        self.vol_slider.setFixedWidth(130)
        self.vol_slider.setFixedHeight(38)
        self.vol_slider.setToolTip("Ana Ses Seviyesi")
        self.vol_slider.valueChanged.connect(self._on_vol_slider)
        ctrl_layout.addWidget(self.vol_slider)

        # Volume percentage label
        self.lbl_vol = QLabel("50%")
        self.lbl_vol.setObjectName("vol_label")
        self.lbl_vol.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        ctrl_layout.addWidget(self.lbl_vol)

        # ── Volume Up ─────────────────────────────────────────────────
        btn_vol_up = QPushButton()
        btn_vol_up.setIcon(_mk_icon('vol_up', 34))
        btn_vol_up.setIconSize(QSize(34, 34))
        btn_vol_up.setObjectName("media_btn")
        btn_vol_up.setToolTip("Ses Artır  [+]")
        btn_vol_up.setFixedSize(52, 44)
        btn_vol_up.clicked.connect(self._vol_up)
        ctrl_layout.addWidget(btn_vol_up)
        self.media_btn_refs.append(btn_vol_up)

        ctrl_layout.addStretch()

        # ── Band buttons ──────────────────────────────────────────────
        for b in [8, 16, 32, 64]:
            btn = QPushButton(f"{b} Band")
            btn.setFixedHeight(28)
            btn.clicked.connect(lambda checked, bb=b: self.set_bands(bb))
            ctrl_layout.addWidget(btn)

        layout.addWidget(ctrl_bar)

        # ── Mid: spectrum + db meter ──────────────────────────────────
        mid = QHBoxLayout()
        mid.setSpacing(4)

        pg.setConfigOptions(antialias=False)
        self.plot_widget = pg.PlotWidget(background='#0D0F17')
        self.plot_widget.setMenuEnabled(False)
        self.plot_widget.setMouseEnabled(x=False, y=False)
        self.plot_widget.hideAxis('left')

        # Lock Y range — disable all auto-ranging
        vb = self.plot_widget.getViewBox()
        vb.setAutoVisible(y=False)
        vb.enableAutoRange(axis='y', enable=False)
        vb.enableAutoRange(axis='x', enable=False)
        vb.setLimits(yMin=0, yMax=100, minYRange=100, maxYRange=100)
        self.plot_widget.setYRange(0, 100, padding=0)

        self.bottom_axis = self.plot_widget.getAxis('bottom')
        self.bottom_axis.setPen(pg.mkPen(color='#2E3345', width=1))
        self.bottom_axis.setTextPen(pg.mkPen(color='#6A7290'))

        self.plot_widget.showAxis('right')
        right_ax = self.plot_widget.getAxis('right')
        right_ax.setPen(pg.mkPen(color='#0D0F17'))
        right_ax.setTextPen(pg.mkPen(color='#5A6280'))
        right_ax.setTicks([[(25,'%25'), (50,'%50'), (75,'%75'), (100,'MAX')]])

        # Reference dashed lines — above scanlines (Z=15)
        for y, col in [(100,'#4A5060'), (75,'#353D4A'), (50,'#353D4A'), (25,'#353D4A')]:
            ln = pg.InfiniteLine(pos=y, angle=0,
                pen=pg.mkPen(color=col, width=1, style=Qt.PenStyle.DashLine))
            ln.setZValue(15)
            self.plot_widget.addItem(ln)

        mid.addWidget(self.plot_widget, stretch=1)

        # DB Meter
        self.db_widget = pg.PlotWidget(background='#13161F')
        self.db_widget.setFixedWidth(80)   # 2× genişlik
        self.db_widget.setMenuEnabled(False)
        self.db_widget.setMouseEnabled(x=False, y=False)
        self.db_widget.hideAxis('bottom')
        self.db_widget.hideAxis('left')

        db_vb = self.db_widget.getViewBox()
        db_vb.setAutoVisible(y=False)
        db_vb.enableAutoRange(axis='y', enable=False)
        db_vb.enableAutoRange(axis='x', enable=False)
        db_vb.setLimits(yMin=0, yMax=100, minYRange=100, maxYRange=100)
        self.db_widget.setXRange(0, 1, padding=0)
        self.db_widget.setYRange(0, 100, padding=0)

        self.db_bar_green  = pg.QtWidgets.QGraphicsRectItem(0.1, 0,  0.8, 0)
        self.db_bar_yellow = pg.QtWidgets.QGraphicsRectItem(0.1, 75, 0.8, 0)
        self.db_bar_red    = pg.QtWidgets.QGraphicsRectItem(0.1, 90, 0.8, 0)
        self.db_bar_green.setBrush(pg.mkBrush(QColor(50, 230, 130)));  self.db_bar_green.setPen(pg.mkPen(None))
        self.db_bar_yellow.setBrush(pg.mkBrush(QColor(255, 175, 40))); self.db_bar_yellow.setPen(pg.mkPen(None))
        self.db_bar_red.setBrush(pg.mkBrush(QColor(255, 55, 55)));     self.db_bar_red.setPen(pg.mkPen(None))
        for bar in (self.db_bar_green, self.db_bar_yellow, self.db_bar_red):
            self.db_widget.addItem(bar)

        self.db_peak_item = pg.QtWidgets.QGraphicsRectItem(0.1, 0, 0.8, self.led_step * 0.75)
        self.db_peak_item.setBrush(self.PEAK_BRUSH)
        self.db_peak_item.setPen(pg.mkPen(None))
        self.db_widget.addItem(self.db_peak_item)

        # LED scanlines
        for y in np.arange(0, 102, self.led_step):
            sl = pg.InfiniteLine(pos=y, angle=0, pen=pg.mkPen('#0D0F17', width=3))
            sl.setZValue(10); self.plot_widget.addItem(sl)
            dl = pg.InfiniteLine(pos=y, angle=0, pen=pg.mkPen('#13161F', width=3))
            dl.setZValue(10); self.db_widget.addItem(dl)

        mid.addWidget(self.db_widget)
        layout.addLayout(mid, stretch=1)

        # Status bar
        bot = QHBoxLayout()
        self.lbl_status = QLabel("")
        self.lbl_status.setObjectName("status")
        bot.addWidget(self.lbl_status)
        bot.addStretch()
        layout.addLayout(bot)

    # ─────────────────────────────────────────────────────────────────
    def _on_vol_slider(self, value):
        """Slider moved → set system volume + update label."""
        self.lbl_vol.setText(f"{value}%")
        if PYCAW_AVAILABLE:
            set_master_volume(value)

    def _vol_down(self):
        if PYCAW_AVAILABLE:
            v = max(0, get_master_volume() - 2)
            set_master_volume(v)
            self.vol_slider.blockSignals(True)
            self.vol_slider.setValue(v)
            self.vol_slider.blockSignals(False)
            self.lbl_vol.setText(f"{v}%")
        else:
            send_media_key(VK_VOLUME_DOWN)

    def _vol_up(self):
        if PYCAW_AVAILABLE:
            v = min(100, get_master_volume() + 2)
            set_master_volume(v)
            self.vol_slider.blockSignals(True)
            self.vol_slider.setValue(v)
            self.vol_slider.blockSignals(False)
            self.lbl_vol.setText(f"{v}%")
        else:
            send_media_key(VK_VOLUME_UP)

    # ─────────────────────────────────────────────────────────────────
    def setup_shortcuts(self):
        sc_map = [
            (Qt.Key.Key_Left,  lambda: send_media_key(VK_MEDIA_PREV_TRACK)),
            (Qt.Key.Key_Right, lambda: send_media_key(VK_MEDIA_NEXT_TRACK)),
            (Qt.Key.Key_Space, lambda: send_media_key(VK_MEDIA_PLAY_PAUSE)),
            (Qt.Key.Key_S,     lambda: send_media_key(VK_MEDIA_STOP)),
            (Qt.Key.Key_M,     lambda: send_media_key(VK_VOLUME_MUTE)),
            (Qt.Key.Key_Minus, self._vol_down),
            (Qt.Key.Key_Plus,  self._vol_up),
            (Qt.Key.Key_Equal, self._vol_up),
        ]
        for key, fn in sc_map:
            sc = QShortcut(QKeySequence(key), self)
            sc.activated.connect(fn)

    # ─────────────────────────────────────────────────────────────────
    def set_bands(self, num_bands):
        self.num_bands  = num_bands
        if hasattr(self, 'settings'):
            self.settings.setValue("band_count", num_bands)
        max_freq        = min(24000, self.fs / 2)
        self.band_edges = np.logspace(np.log10(20), np.log10(max_freq), num_bands + 1)
        self.band_centers = (self.band_edges[:-1] + self.band_edges[1:]) / 2

        self.smoothed_bands   = np.zeros(num_bands)
        self.peaks            = np.zeros(num_bands)
        self.peak_hold_frames = np.zeros(num_bands)
        self.peak_hold_max    = 30
        self.peak_drop_rate   = 1.0

        if hasattr(self, 'bar_items'):
            for item in getattr(self, 'bar_items', []): self.plot_widget.removeItem(item)
        for item in getattr(self, 'bar_items_cyan', []):   self.plot_widget.removeItem(item)
        for item in getattr(self, 'bar_items_yellow', []): self.plot_widget.removeItem(item)
        for item in getattr(self, 'bar_items_red', []):    self.plot_widget.removeItem(item)
        for item in getattr(self, 'peak_items', []):       self.plot_widget.removeItem(item)

        self.bar_items_cyan   = []
        self.bar_items_yellow = []
        self.bar_items_red    = []
        self.peak_items       = []

        for i in range(num_bands):
            rc = pg.QtWidgets.QGraphicsRectItem(i - 0.4, 0, 0.8, 0)
            rc.setPen(pg.mkPen(None)); rc.setBrush(self.BAR_BRUSH)
            self.plot_widget.addItem(rc); self.bar_items_cyan.append(rc)

            ry = pg.QtWidgets.QGraphicsRectItem(i - 0.4, 75, 0.8, 0)
            ry.setPen(pg.mkPen(None)); ry.setBrush(self.DB_PEAK_YELLOW)
            self.plot_widget.addItem(ry); self.bar_items_yellow.append(ry)

            rr = pg.QtWidgets.QGraphicsRectItem(i - 0.4, 90, 0.8, 0)
            rr.setPen(pg.mkPen(None)); rr.setBrush(self.DB_PEAK_RED)
            self.plot_widget.addItem(rr); self.bar_items_red.append(rr)

            p = pg.QtWidgets.QGraphicsRectItem(i - 0.4, 0, 0.8, self.led_step * 0.75)
            p.setPen(pg.mkPen(None)); p.setBrush(self.BAR_BRUSH)
            self.plot_widget.addItem(p); self.peak_items.append(p)

        # Re-lock ranges after band change (prevents auto-range from kicking in)
        x0 = -0.8
        x1 = num_bands - 0.2
        vb = self.plot_widget.getViewBox()
        vb.setLimits(xMin=x0, xMax=x1, yMin=0, yMax=100,
                     minXRange=x1-x0, maxXRange=x1-x0,
                     minYRange=100,   maxYRange=100)
        self.plot_widget.setXRange(x0, x1, padding=0)
        self.plot_widget.setYRange(0, 100, padding=0)

        # Axis ticks
        labels_8  = ["50Hz","125Hz","250Hz","1kHz","2kHz","4kHz","8kHz","16kHz"]
        labels_16 = ["20Hz","63Hz","100Hz","160Hz","250Hz","400Hz","630Hz","1kHz",
                     "1.6kHz","2.5kHz","4kHz","8kHz","10kHz","12.5kHz","16kHz","18kHz"]
        ticks = []
        if   num_bands ==  8: ticks = [(i,   labels_8[i])  for i in range(8)]
        elif num_bands == 16: ticks = [(i,   labels_16[i]) for i in range(16)]
        elif num_bands == 32: ticks = [(i*2, labels_16[i]) for i in range(16)]
        elif num_bands == 64: ticks = [(i*4, labels_16[i]) for i in range(16)]
        else:
            step = max(1, num_bands // 8)
            for i in range(0, num_bands, step):
                f = int(self.band_centers[i])
                ticks.append((i, f"{f/1000:.1f}k" if f >= 1000 else str(f)))
        self.bottom_axis.setTicks([ticks])

    # ─────────────────────────────────────────────────────────────────
    def setup_audio(self):
        try:
            self.p = pa.PyAudio()
            self.p.get_host_api_info_by_type(pa.paWASAPI)
            spk = self.p.get_default_wasapi_loopback()

            self.fs       = int(spk["defaultSampleRate"])
            self.fft_size = 8192
            self.audio_data = np.zeros(self.fft_size)

            self.stream = self.p.open(
                format=pa.paFloat32,
                channels=spk["maxInputChannels"],
                rate=self.fs,
                frames_per_buffer=self.fft_size // 8,
                input=True,
                input_device_index=spk["index"],
                stream_callback=self.audio_callback
            )
            self.stream.start_stream()

            name = spk.get('name', 'Ses Cihazı')
            if len(name) > 42: name = name[:42] + "…"
            self.lbl_status.setText(f"🎧  {name}")

            if self.band_edges is not None:
                self.set_bands(self.num_bands)

        except Exception:
            try:
                with open("error.txt", "a", encoding="utf-8") as f:
                    import traceback; f.write(traceback.format_exc())
            except Exception:
                pass

    # ─────────────────────────────────────────────────────────────────
    def audio_callback(self, in_data, frame_count, time_info, status):
        if in_data and not self.audio_lock:
            raw  = np.frombuffer(in_data, dtype=np.float32)
            chs  = self.stream._channels
            mono = raw.reshape(-1, chs).mean(axis=1) if chs > 1 else raw.copy()
            n = len(mono)
            if n >= self.fft_size:
                self.audio_data = mono[-self.fft_size:]
            else:
                self.audio_data = np.roll(self.audio_data, -n)
                self.audio_data[-n:] = mono
        return (in_data, pa.paContinue)

    # ─────────────────────────────────────────────────────────────────
    def update_ui(self):
        # ── CPU Tasarrufu (Idle Detection) ───────────────────────────
        # Eğer ses yoksa ve UI zaten sıfırlanmışsa, ağır FFT ve UI güncellemelerini atla
        recent_audio = self.audio_data[-1024:]
        rms = np.sqrt(np.mean(recent_audio ** 2) + 1e-12)
        
        is_ui_zero = (np.max(self.smoothed_bands) < 0.1 and np.max(self.peaks) < 0.1 and 
                      self.smoothed_db < 0.1 and self.db_peak < 0.1)
                      
        if rms < 1e-4 and is_ui_zero:
            # Sadece ses senkronizasyonunu yap
            if PYCAW_AVAILABLE:
                self._vol_sync_counter += 1
                if self._vol_sync_counter >= 100:
                    self._vol_sync_counter = 0
                    real_vol = get_master_volume()
                    if abs(real_vol - self.vol_slider.value()) > 2:
                        self.vol_slider.blockSignals(True)
                        self.vol_slider.setValue(real_vol)
                        self.vol_slider.blockSignals(False)
                        self.lbl_vol.setText(f"{real_vol}%")
            return

        # ── Spectrum (WinAmp / JetAudio absolute-dB style) ───────────
        window   = np.hanning(self.fft_size)
        windowed = self.audio_data * window
        fft_raw  = np.fft.rfft(windowed)

        w_sum   = np.sum(window) + 1e-30
        fft_mag = np.abs(fft_raw) * 2.0 / w_sum

        freq_res     = self.fs / self.fft_size
        band_db_vals = np.full(self.num_bands, -120.0)

        for i in range(self.num_bands):
            s = max(1, int(self.band_edges[i]     / freq_res))
            e = max(s + 1, int(self.band_edges[i+1] / freq_res))
            e = min(e, len(fft_mag))
            if s < e:
                band_db_vals[i] = 20.0 * np.log10(np.max(fft_mag[s:e]) + 1e-12)

        DB_FLOOR = -49.0
        DB_CEIL  =  -4.0
        DB_RANGE = DB_CEIL - DB_FLOOR
        band_values = np.clip(
            (band_db_vals - DB_FLOOR) / DB_RANGE * 100.0, 0.0, 100.0
        )

        # Fast attack / moderate decay
        RISE = 0.92
        FALL = 0.28
        mask = band_values > self.smoothed_bands
        self.smoothed_bands = np.where(
            mask,
            self.smoothed_bands + RISE * (band_values - self.smoothed_bands),
            self.smoothed_bands + FALL * (band_values - self.smoothed_bands)
        )

        # Peak hold
        for i in range(self.num_bands):
            v = self.smoothed_bands[i]
            if v >= self.peaks[i]:
                self.peaks[i] = v
                self.peak_hold_frames[i] = self.peak_hold_max
            else:
                if self.peak_hold_frames[i] > 0:
                    self.peak_hold_frames[i] -= 1
                else:
                    self.peaks[i] = max(self.peaks[i] - self.peak_drop_rate, v, 0.0)

        # Draw bars and update colors
        led = self.led_step
        for i in range(self.num_bands):
            h  = int(self.smoothed_bands[i] / led) * led
            pv = int(self.peaks[i]          / led) * led
            
            self.bar_items_cyan[i].setRect(  i - 0.4,  0, 0.8, min(75, h))
            self.bar_items_yellow[i].setRect(i - 0.4, 75, 0.8, max(0, min(15, h - 75)))
            self.bar_items_red[i].setRect(   i - 0.4, 90, 0.8, max(0, min(10, h - 90)))
            
            self.peak_items[i].setRect(i - 0.4, pv, 0.8, led * 0.75)

            if pv > 90:   self.peak_items[i].setBrush(self.DB_PEAK_RED)
            elif pv > 75: self.peak_items[i].setBrush(self.DB_PEAK_YELLOW)
            else:         self.peak_items[i].setBrush(self.BAR_BRUSH)

        # ── DB Meter (RMS / volume-sensitive) ────────────────────────
        db_pct = float(np.clip((20.0 * np.log10(rms) + 60.0) * (100.0 / 60.0), 0.0, 100.0))

        if db_pct > self.smoothed_db:
            self.smoothed_db = self.smoothed_db * 0.05 + db_pct * 0.95
        else:
            self.smoothed_db = self.smoothed_db * 0.50 + db_pct * 0.50

        dbh = int(self.smoothed_db / led) * led
        self.db_bar_green.setRect( 0.1,  0,  0.8, min(75, dbh))
        self.db_bar_yellow.setRect(0.1, 75, 0.8, max(0, min(15, dbh - 75)))
        self.db_bar_red.setRect(   0.1, 90, 0.8, max(0, min(10, dbh - 90)))

        if self.smoothed_db >= self.db_peak:
            self.db_peak = self.smoothed_db
            self.db_peak_hold_frames = self.peak_hold_max
        else:
            if self.db_peak_hold_frames > 0:
                self.db_peak_hold_frames -= 1
            else:
                self.db_peak = max(self.db_peak - self.peak_drop_rate, self.smoothed_db, 0.0)

        db_pv = int(self.db_peak / led) * led
        self.db_peak_item.setRect(0.1, db_pv, 0.8, led * 0.75)

        if self.db_peak <= 75:
            self.db_peak_item.setBrush(self.DB_PEAK_GREEN)
        elif self.db_peak <= 90:
            self.db_peak_item.setBrush(self.DB_PEAK_YELLOW)
        else:
            self.db_peak_item.setBrush(self.DB_PEAK_RED)

        # ── Sync slider with real volume every ~2 s (100 frames) ─────
        if PYCAW_AVAILABLE:
            self._vol_sync_counter += 1
            if self._vol_sync_counter >= 100:
                self._vol_sync_counter = 0
                real_vol = get_master_volume()
                if abs(real_vol - self.vol_slider.value()) > 2:
                    self.vol_slider.blockSignals(True)
                    self.vol_slider.setValue(real_vol)
                    self.vol_slider.blockSignals(False)
                    self.lbl_vol.setText(f"{real_vol}%")

    # ─────────────────────────────────────────────────────────────────
    def update_media_labels(self, title, artist):
        if title != self.current_title or artist != self.current_artist:
            self.current_title  = title
            self.current_artist = artist
            self.lbl_title.setText(title)
            self.lbl_artist.setText(artist)

    # ─────────────────────────────────────────────────────────────────
    def show_info(self):
        text = (
            "31.5 Hz: Sub-sonik (Hissedilen bas)\n\n"
            "63 Hz:   Derin bas\n\n"
            "100 Hz:  Bas vuruşu (Punch)\n\n"
            "160 Hz:  Sıcaklık (Warmth)\n\n"
            "250 Hz:  Alt mid\n\n"
            "400 Hz:  Gövde / Tokluk\n\n"
            "630 Hz:  Burun sesi bölgesi\n\n"
            "1 kHz:   Odak noktası\n\n"
            "1.6 kHz: Projeksiyon\n\n"
            "2.5 kHz: Netlik ve saldırı (Attack)\n\n"
            "4 kHz:   Tanımlama (Definition)\n\n"
            "6.3 kHz: Detay\n\n"
            "8 kHz:   Tiz parlaklığı\n\n"
            "10 kHz:  Hava (Air)\n\n"
            "12.5 kHz: Sibilans yönetimi\n\n"
            "16 kHz:  Ultra-tiz (Brilliance)"
        )
        msg = QMessageBox(self)
        msg.setWindowTitle("Frekans Bilgileri")
        msg.setText(text)
        msg.setStyleSheet("""
            QMessageBox { background-color: #13161F; }
            QLabel { color:#FFFFFF; font-size:13px; font-family:'Segoe UI'; }
            QPushButton { background-color:#1E2130; color:#FFFFFF; padding:5px 15px;
                          border-radius:4px; font-weight:bold; border:1px solid #2A2F40; }
            QPushButton:hover { background-color:#262C42; }
        """)
        msg.exec()

    # ─────────────────────────────────────────────────────────────────
    def closeEvent(self, event):
        self.update_timer.stop()
        if WINSDK_AVAILABLE:
            self.media_thread.requestInterruption()
            self.media_thread.wait(2000)
        self.audio_lock = True
        if self.stream:
            try: self.stream.stop_stream(); self.stream.close()
            except Exception: pass
        if self.p:
            try: self.p.terminate()
            except Exception: pass
        event.accept()


# ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    try:
        app.setFont(QFont("Segoe UI", 10))
        window = RealTimeSpekApp()
        window.show()
        sys.exit(app.exec())
    except Exception:
        import traceback
        tb = traceback.format_exc()
        try:
            with open("error.txt", "a", encoding="utf-8") as f: f.write(tb)
        except Exception:
            pass
        try:
            msg = QMessageBox()
            msg.setWindowTitle("Başlatma Hatası")
            msg.setText(f"Uygulama başlatılamadı:\n\n{tb[-900:]}")
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.exec()
        except Exception:
            pass
