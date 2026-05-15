import sys
from PyQt6.QtWidgets import QApplication

app = QApplication(sys.argv)

import asyncio
import numpy as np
from PyQt6.QtWidgets import (QMainWindow, QVBoxLayout, 
                             QWidget, QLabel, QHBoxLayout, QPushButton, QMessageBox)
from PyQt6.QtCore import QTimer, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QColor, QBrush
import pyqtgraph as pg

import pyaudiowpatch as pa

try:
    from winsdk.windows.media.control import GlobalSystemMediaTransportControlsSessionManager
    WINSDK_AVAILABLE = True
except ImportError:
    WINSDK_AVAILABLE = False

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
            # Check for interruption more frequently
            for _ in range(20):
                if self.isInterruptionRequested():
                    break
                self.msleep(100)
        loop.close()
            
    async def fetch(self):
        manager = await GlobalSystemMediaTransportControlsSessionManager.request_async()
        current_session = manager.get_current_session()
        if current_session:
            info = await current_session.try_get_media_properties_async()
            title = info.title if info.title else "Bilinmeyen Şarkı"
            artist = info.artist if info.artist else "Bilinmeyen Sanatçı"
            return (title, artist)
        return ("Müzik Çalmıyor", "-")

class RealTimeSpekApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Real-Time Audio Quality Analyzer")
        self.resize(850, 500)
        
        # Audio
        self.p = None
        self.stream = None
        self.fs = 48000
        
        # FFT Buffer
        self.fft_size = 4096
        self.freqs = np.fft.rfftfreq(self.fft_size, 1 / self.fs)
        self.audio_data = np.zeros(self.fft_size)
        
        # State
        self.current_title = "Bekleniyor..."
        self.current_artist = "Sanatçı bilgisi yok"
        
        self.init_ui()
        self.setup_audio()
        
        # Default 32 Bands
        self.set_bands(32)
        
        # Timers
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_ui)
        self.update_timer.start(30)
        
        if WINSDK_AVAILABLE:
            self.media_thread = MediaInfoThread()
            self.media_thread.info_ready.connect(self.update_media_labels)
            self.media_thread.start()

    def init_ui(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #0F111A;
            }
            QLabel#title {
                color: #FFFFFF;
                font-size: 24px;
                font-weight: bold;
                font-family: 'Segoe UI', sans-serif;
            }
            QLabel#artist {
                color: #8A91A6;
                font-size: 14px;
                font-family: 'Segoe UI', sans-serif;
            }
            QLabel#status {
                color: #5C6370;
                font-size: 12px;
            }
            QWidget#top_bar {
                background-color: #161925;
                border-bottom: 1px solid #282C34;
                border-radius: 6px;
            }
            QPushButton {
                background-color: #282C34;
                color: #FFFFFF;
                padding: 6px 12px;
                border-radius: 4px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background-color: #3E4452;
            }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(15)

        # Top Bar
        top_bar = QWidget()
        top_bar.setObjectName("top_bar")
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(10, 10, 10, 10)
        
        left_placeholder = QWidget()
        left_placeholder.setFixedSize(30, 30)
        top_layout.addWidget(left_placeholder)
        
        center_widget = QWidget()
        center_layout = QVBoxLayout(center_widget)
        center_layout.setContentsMargins(0, 0, 0, 0)
        
        self.lbl_title = QLabel(self.current_title)
        self.lbl_title.setObjectName("title")
        self.lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.lbl_artist = QLabel(self.current_artist)
        self.lbl_artist.setObjectName("artist")
        self.lbl_artist.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        center_layout.addWidget(self.lbl_title)
        center_layout.addWidget(self.lbl_artist)
        top_layout.addWidget(center_widget, stretch=1)
        
        self.btn_info = QPushButton("?")
        self.btn_info.setFixedSize(30, 30)
        self.btn_info.clicked.connect(self.show_info)
        top_layout.addWidget(self.btn_info)
        
        layout.addWidget(top_bar)

        # Middle Section (Plot + DB Meter)
        mid_layout = QHBoxLayout()
        
        # Main Plot
        pg.setConfigOptions(antialias=True)
        self.plot_widget = pg.PlotWidget(background='#0F111A')
        self.plot_widget.setMenuEnabled(False)
        self.plot_widget.setMouseEnabled(x=False, y=False)
        self.plot_widget.hideAxis('left')
        self.plot_widget.setYRange(0, 100)
        
        self.bottom_axis = self.plot_widget.getAxis('bottom')
        self.bottom_axis.setPen(pg.mkPen(color='#3E4452', width=1))
        self.bottom_axis.setTextPen(pg.mkPen(color='#8A91A6'))
        
        self.plot_widget.showAxis('right')
        right_axis = self.plot_widget.getAxis('right')
        right_axis.setPen(pg.mkPen(color='#0F111A'))
        right_axis.setTextPen(pg.mkPen(color='#8A91A6'))
        right_axis.setTicks([[(25, '%25'), (50, '%50'), (75, '%75'), (100, 'MAX')]])
        
        for y, color, style in [
            (100, '#5C6370', Qt.PenStyle.SolidLine),
            (75, '#3E4452', Qt.PenStyle.DashLine),
            (50, '#5C6370', Qt.PenStyle.DashLine),
            (25, '#3E4452', Qt.PenStyle.DashLine)
        ]:
            line = pg.InfiniteLine(pos=y, angle=0, pen=pg.mkPen(color=color, width=1, style=style))
            line.setZValue(-1)
            self.plot_widget.addItem(line)
        
        mid_layout.addWidget(self.plot_widget, stretch=1)
        
        # DB Meter Plot
        self.db_widget = pg.PlotWidget(background='#161925')
        self.db_widget.setFixedWidth(40)
        self.db_widget.setMenuEnabled(False)
        self.db_widget.setMouseEnabled(x=False, y=False)
        self.db_widget.hideAxis('bottom')
        self.db_widget.hideAxis('left')
        
        self.db_widget.setXRange(0, 1)
        self.db_widget.setYRange(0, 100)
        
        self.db_bar_green = pg.QtWidgets.QGraphicsRectItem(0.1, 0, 0.8, 0)
        self.db_bar_green.setBrush(pg.mkBrush(QColor(60, 255, 120)))
        self.db_bar_green.setPen(pg.mkPen(None))
        self.db_widget.addItem(self.db_bar_green)
        
        self.db_bar_yellow = pg.QtWidgets.QGraphicsRectItem(0.1, 75, 0.8, 0)
        self.db_bar_yellow.setBrush(pg.mkBrush(QColor(255, 180, 50)))
        self.db_bar_yellow.setPen(pg.mkPen(None))
        self.db_widget.addItem(self.db_bar_yellow)
        
        self.db_bar_red = pg.QtWidgets.QGraphicsRectItem(0.1, 90, 0.8, 0)
        self.db_bar_red.setBrush(pg.mkBrush(QColor(255, 60, 60)))
        self.db_bar_red.setPen(pg.mkPen(None))
        self.db_widget.addItem(self.db_bar_red)
        
        # --- LED MATRIX EFFECT (SCANLINES) ---
        self.led_step = 2.0
        for y in np.arange(0, 105, self.led_step):
            line = pg.InfiniteLine(pos=y, angle=0, pen=pg.mkPen('#0F111A', width=3))
            line.setZValue(10)
            self.plot_widget.addItem(line)
            
            db_line = pg.InfiniteLine(pos=y, angle=0, pen=pg.mkPen('#161925', width=3))
            db_line.setZValue(10)
            self.db_widget.addItem(db_line)
        # -------------------------------------
        
        mid_layout.addWidget(self.db_widget)
        
        layout.addLayout(mid_layout, stretch=1)
        
        # Bottom Controls
        bottom_layout = QHBoxLayout()
        self.lbl_status = QLabel("Ses cihazı bekleniyor...")
        self.lbl_status.setObjectName("status")
        bottom_layout.addWidget(self.lbl_status)
        
        bottom_layout.addStretch()
        
        # Band Selection Buttons
        for bands in [8, 16, 32, 64]:
            btn = QPushButton(f"{bands} Band")
            btn.clicked.connect(lambda checked, b=bands: self.set_bands(b))
            bottom_layout.addWidget(btn)
            
        layout.addLayout(bottom_layout)
        
        self.bar_items = []
        self.peak_items = []

    def set_bands(self, num_bands):
        self.num_bands = num_bands
        max_freq = min(24000, self.fs / 2)
        # Re-create edges (Logarithmic)
        self.band_edges = np.logspace(np.log10(20), np.log10(max_freq), self.num_bands + 1)
        self.band_centers = (self.band_edges[:-1] + self.band_edges[1:]) / 2
        
        self.smoothed_bands = np.zeros(self.num_bands)
        self.peaks = np.zeros(self.num_bands)
        self.peak_hold_frames = np.zeros(self.num_bands)
        self.peak_hold_max = 20
        self.peak_drop_rate = 3.0
        
        for item in self.bar_items:
            self.plot_widget.removeItem(item)
        for item in self.peak_items:
            self.plot_widget.removeItem(item)
            
        self.bar_items = []
        self.peak_items = []
        
        for i in range(self.num_bands):
            rect = pg.QtWidgets.QGraphicsRectItem(i - 0.4, 0, 0.8, 0)
            rect.setPen(pg.mkPen(None))
            self.plot_widget.addItem(rect)
            self.bar_items.append(rect)
            
            prect = pg.QtWidgets.QGraphicsRectItem(i - 0.4, 0, 0.8, 1.5)
            prect.setBrush(pg.mkBrush('#FFFFFF'))
            prect.setPen(pg.mkPen(None))
            self.plot_widget.addItem(prect)
            self.peak_items.append(prect)
            
        self.plot_widget.setXRange(-1, self.num_bands)
        
        # Update Ticks
        ticks = []
        
        labels_8 = ["50Hz", "125Hz", "250Hz", "1kHz", "2kHz", "4kHz", "8kHz", "16kHz"]
        labels_16 = ["20Hz", "63Hz", "100Hz", "160Hz", "250Hz", "400Hz", "630Hz", "1kHz", 
                     "1.6kHz", "2.5kHz", "4kHz", "8kHz", "10kHz", "12.5kHz", "16kHz", "18kHz"]
                     
        if self.num_bands == 8:
            for i in range(8):
                ticks.append((i, labels_8[i]))
        elif self.num_bands == 16:
            for i in range(16):
                ticks.append((i, labels_16[i]))
        elif self.num_bands == 32:
            step = 32 // 16
            for i in range(16):
                ticks.append((i * step, labels_16[i]))
        elif self.num_bands == 64:
            step = 64 // 16
            for i in range(16):
                ticks.append((i * step, labels_16[i]))
        else:
            step = max(1, self.num_bands // 8)
            for i in range(0, self.num_bands, step):
                freq = int(self.band_centers[i])
                label = f"{freq/1000:.1f}k" if freq >= 1000 else str(freq)
                ticks.append((i, label))
                
            if len(ticks) > 0 and ticks[-1][0] < self.num_bands - 1:
                freq = int(self.band_centers[-1])
                label = f"{freq/1000:.1f}k" if freq >= 1000 else str(freq)
                ticks.append((self.num_bands - 1, label))
            
        self.bottom_axis.setTicks([ticks])

    def setup_audio(self):
        try:
            self.p = pa.PyAudio()
            wasapi_info = self.p.get_host_api_info_by_type(pa.paWASAPI)
            default_speakers = self.p.get_default_wasapi_loopback()
            
            self.fs = int(default_speakers["defaultSampleRate"])
            self.fft_size = 4096
            self.freqs = np.fft.rfftfreq(self.fft_size, 1 / self.fs)
            self.audio_data = np.zeros(self.fft_size)
            
            self.lbl_status.setText(f"Dinleniyor: {default_speakers['name']}")
            
            self.stream = self.p.open(format=pa.paFloat32,
                            channels=default_speakers["maxInputChannels"],
                            rate=self.fs,
                            frames_per_buffer=self.fft_size // 4,
                            input=True,
                            input_device_index=default_speakers["index"],
                            stream_callback=self.audio_callback)
            self.stream.start_stream()
            
            # Re-initialize bands with correct sample rate
            self.set_bands(self.num_bands)
            
        except Exception as e:
            self.lbl_status.setText(f"Ses cihazı hatası: {e}")

    def audio_callback(self, in_data, frame_count, time_info, status):
        if in_data:
            audio_data = np.frombuffer(in_data, dtype=np.float32)
            channels = self.stream._channels
            if channels > 1:
                mono = audio_data.reshape(-1, channels).mean(axis=1)
            else:
                mono = audio_data
            
            if len(mono) >= self.fft_size:
                self.audio_data = mono[-self.fft_size:]
            else:
                self.audio_data = np.roll(self.audio_data, -len(mono))
                self.audio_data[-len(mono):] = mono
                
        return (in_data, pa.paContinue)

    def update_ui(self):
        # Apply Hanning window
        window = np.hanning(len(self.audio_data))
        windowed_data = self.audio_data * window
        
        # Calculate FFT properly normalized
        fft_data = np.abs(np.fft.rfft(windowed_data)) * 2.0 / self.fft_size
        fft_db = 20 * np.log10(fft_data + 1e-12)
        # Optimize scaling so it reaches 100 at typical max music volume
        fft_db = np.clip((fft_db + 65) * 2.2, 0, 100)
        
        band_values = np.zeros(self.num_bands)
        for i in range(self.num_bands):
            start_freq = self.band_edges[i]
            end_freq = self.band_edges[i+1]
            
            start_idx = int(start_freq / (self.fs / self.fft_size))
            end_idx = int(end_freq / (self.fs / self.fft_size))
            
            if start_idx == end_idx:
                end_idx += 1
                
            start_idx = max(0, min(start_idx, len(fft_db)-1))
            end_idx = max(1, min(end_idx, len(fft_db)))
            
            if start_idx < end_idx:
                band_values[i] = np.max(fft_db[start_idx:end_idx])
            
        self.smoothed_bands = self.smoothed_bands * 0.5 + band_values * 0.5
        
        for i in range(self.num_bands):
            val = self.smoothed_bands[i]
            if val >= self.peaks[i]:
                self.peaks[i] = val
                self.peak_hold_frames[i] = self.peak_hold_max
            else:
                if self.peak_hold_frames[i] > 0:
                    self.peak_hold_frames[i] -= 1
                else:
                    self.peaks[i] -= self.peak_drop_rate
                    if self.peaks[i] < val:
                        self.peaks[i] = val
                    if self.peaks[i] < 0:
                        self.peaks[i] = 0

        base_hue = 0.55
        for i in range(self.num_bands):
            # Snap to LED steps
            h = np.floor(self.smoothed_bands[i] / self.led_step) * self.led_step
            self.bar_items[i].setRect(i - 0.4, 0, 0.8, h)
            
            brightness = min(1.0, max(0.3, h / 80.0 + 0.3))
            color = QColor()
            color.setHsvF(base_hue, 1.0, brightness)
            self.bar_items[i].setBrush(pg.mkBrush(color))
            
            p_val = np.floor(self.peaks[i] / self.led_step) * self.led_step
            self.peak_items[i].setRect(i - 0.4, p_val, 0.8, self.led_step * 0.8)

        # Update DB Meter
        rms = np.sqrt(np.mean(self.audio_data**2) + 1e-12)
        db = 20 * np.log10(rms)
        db_percent = np.clip((db + 55) * (100 / 55), 0, 100)
        
        if not hasattr(self, 'smoothed_db'):
            self.smoothed_db = db_percent
        else:
            self.smoothed_db = self.smoothed_db * 0.6 + db_percent * 0.4
            
        db_height = np.floor(self.smoothed_db / self.led_step) * self.led_step
        
        h_green = min(75, db_height)
        self.db_bar_green.setRect(0.1, 0, 0.8, h_green)
        
        h_yellow = max(0, min(15, db_height - 75))
        self.db_bar_yellow.setRect(0.1, 75, 0.8, h_yellow)
        
        h_red = max(0, min(10, db_height - 90))
        self.db_bar_red.setRect(0.1, 90, 0.8, h_red)

    def update_media_labels(self, title, artist):
        if title != self.current_title or artist != self.current_artist:
            self.current_title = title
            self.current_artist = artist
            self.lbl_title.setText(title)
            self.lbl_artist.setText(artist)

    def show_info(self):
        info_text = (
            "31.5 Hz: Sub-sonik (Hissedilen bas)\n\n"
            "63 Hz: Derin bas\n\n"
            "100 Hz: Bas vuruşu (Punch)\n\n"
            "160 Hz: Sıcaklık (Warmth)\n\n"
            "250 Hz: Alt mid\n\n"
            "400 Hz: Gövde / Tokluk\n\n"
            "630 Hz: Burun sesi bölgesi\n\n"
            "1 kHz: Odak noktası\n\n"
            "1.6 kHz: Projeksiyon\n\n"
            "2.5 kHz: Netlik ve saldırı (Attack)\n\n"
            "4 kHz: Tanımlama (Definition)\n\n"
            "6.3 kHz: Detay\n\n"
            "8 kHz: Tiz parlaklığı\n\n"
            "10 kHz: Hava (Air)\n\n"
            "12.5 kHz: Sibilans yönetimi\n\n"
            "16 kHz: Ultra-tiz (Brilliance)"
        )
        msg = QMessageBox(self)
        msg.setWindowTitle("Frekans Bilgileri")
        msg.setText(info_text)
        msg.setStyleSheet("""
            QMessageBox {
                background-color: #161925;
            }
            QLabel {
                color: #FFFFFF;
                font-size: 13px;
                font-family: 'Segoe UI', sans-serif;
            }
            QPushButton {
                background-color: #282C34;
                color: #FFFFFF;
                padding: 6px 15px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #3E4452;
            }
        """)
        msg.exec()

    def closeEvent(self, event):
        # Stop UI updates
        self.update_timer.stop()
        
        if WINSDK_AVAILABLE:
            self.media_thread.requestInterruption()
            # Wait with a timeout to prevent hanging if thread is stuck
            if not self.media_thread.wait(1000):
                print("Media thread did not terminate in time.")
            
        if self.stream:
            try:
                if self.stream.is_active():
                    self.stream.stop_stream()
                self.stream.close()
            except Exception:
                pass
                
        if self.p:
            try:
                self.p.terminate()
            except Exception:
                pass
                
        event.accept()

if __name__ == "__main__":
    try:
        font = QFont("Segoe UI", 10)
        app.setFont(font)
        
        window = RealTimeSpekApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        import traceback
        with open("error_log.txt", "w") as f:
            f.write(traceback.format_exc())
