# =============================================================================
# FILE: tutorial.py
# PERAN DALAM PROJECT: Layar tutorial & dokumentasi in-app
#
# File ini bertanggung jawab menampilkan layar panduan penggunaan INVENTRA
# yang bisa diakses kapan saja oleh user tanpa harus keluar dari aplikasi.
# Tutorial ini bersifat MODULAR — kontennya diambil dari file gambar eksternal
# (bukan di-hardcode di kode), sehingga mudah diupdate tanpa ubah kode.
#
# STRUKTUR LAYAR:
#   ┌─────────────────────────────────────────────────────┐
#   │  HEADER: "Quick Tutorial"              [✕ Close]    │
#   ├──────────────┬──────────────────────────────────────┤
#   │  SIDEBAR     │  CONTENT AREA (scrollable)           │
#   │  ─────────   │  ───────────────────────────────     │
#   │  ⓘ About    │  (konten berubah tergantung           │
#   │  Tutorial > │   pilihan di sidebar)                 │
#   │  Logika   > │                                       │
#   └──────────────┴──────────────────────────────────────┘
#
# DEPENDENSI FILE EKSTERNAL (harus ada di folder program):
#   - inventraabt.png      → gambar untuk halaman About
#   - tutorial_upload.png  → gambar langkah 1
#   - tutorial_mapping.png → gambar langkah 2
#   - tutorial_preview.png → gambar langkah 3
#   - tutorial_reset.png   → gambar langkah 4
#   - rumus_slide/*.png    → slide-slide penjelasan logika perhitungan
# =============================================================================

import tkinter as tk
from PIL import Image, ImageTk
import sys
import os


# =============================================================================
# HELPER: resource_path
# Fungsi ini menyelesaikan masalah path file saat aplikasi dijalankan
# sebagai .exe (via PyInstaller) vs. saat dijalankan sebagai .py biasa.
#
# MENGAPA PERLU?
#   Saat PyInstaller membuat .exe, semua file (gambar, dll) diekstrak
#   ke folder sementara di sys._MEIPASS (misal: C:\Users\...\AppData\Temp\_MEI12345)
#   Bukan di folder .exe itu sendiri.
#   Tanpa fungsi ini, open("tutorial_upload.png") akan gagal karena file
#   dicari di direktori kerja saat ini, bukan di _MEIPASS.
# =============================================================================
def resource_path(relative_path):
    """Return path yang benar saat jadi EXE atau dijalankan di Python biasa"""
    try:
        # sys._MEIPASS hanya ada saat dijalankan sebagai .exe (PyInstaller)
        base_path = sys._MEIPASS
    except AttributeError:
        # Saat dijalankan sebagai .py biasa → pakai direktori kerja saat ini
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# =============================================================================
# KELAS: TutorialScreen
# Layar tutorial lengkap yang di-inject ke atas main window menggunakan
# .place(relwidth=1, relheight=1) — menutup seluruh area window.
# Saat ditutup (close_tutorial), frame ini di-destroy dan main window kembali terlihat.
#
# POLA DESAIN: "Overlay Screen"
#   TutorialScreen tidak membuat window baru (Toplevel), melainkan frame besar
#   yang menimpa konten main window. Ini memberikan UX yang mulus tanpa
#   dialog terpisah yang bisa kehilangan focus.
# =============================================================================
class TutorialScreen:
    def __init__(self, root, colors, on_continue):
        """
        root        : window Tkinter induk (main window INVENTRA)
        colors      : dict tema warna aplikasi (primary, bg, text, dll)
        on_continue : callback yang dipanggil saat user selesai tutorial
                      (saat ini tidak dipakai karena ada tombol Close)
        """
        self.root        = root
        self.colors      = colors
        self.on_continue = on_continue  # disimpan tapi belum dipakai (future use)
        self.current_step= 0            # untuk navigasi step-by-step (future use)
        
        # ── Definisi konten tutorial ─────────────────────────────────────────
        # Setiap step adalah dict dengan 4 kunci:
        #   title       → judul langkah
        #   description → penjelasan multi-baris
        #   image       → nama file gambar (dicari via resource_path)
        #   tips        → kotak tips berwarna di bawah gambar
        self.tutorial_steps = [
            {
                "title"      : "1. Upload File Excel",
                "description": "Klik tombol '+ Upload File' dan pilih file Excel (.xlsx/.xls)\nAnda bisa pilih beberapa file sekaligus",
                "image"      : "tutorial_upload.png",
                "tips"       : "💡 Pastikan file Excel tidak sedang dibuka di aplikasi lain"
            },
            {
                "title"      : "2. Process & Mapping Data",
                "description": "Klik '⚙ Process', lalu pilih Source File dan Sheet Name untuk setiap tab\nAtur nama kolom sesuai dengan Excel Anda",
                "image"      : "tutorial_mapping.png",
                "tips"       : "⚠️ Nama kolom harus PERSIS sama dengan di Excel!"
            },
            {
                "title"      : "3. Preview & Save",
                "description": "Lihat hasil di Data Preview, lalu klik '💾 Save Excel' untuk menyimpan",
                "image"      : "tutorial_preview.png",
                "tips"       : "✅ Selalu periksa preview sebelum menyimpan"
            },
            {
                "title"      : "4. Reset Data",
                "description": "Klik tombol 'Reset All' untuk membersihkan semua file dan hasil preview.\nGunakan ini jika ingin memulai ulang dari awal.",
                "image"      : "tutorial_reset.png",
                "tips"       : "🗑️ Reset Data akan menghapus semua file yang sudah diupload dan hasil proses."
            }
        ]
        
        self.create_tutorial_screen()

    # =========================================================================
    # HELPER: _darken_color
    # Menggelapkan warna hex sebesar 30 unit per channel RGB.
    # Dipakai untuk efek hover dan klik tombol — memberikan feedback visual.
    # Contoh: "#7F1D1D" → "#610707" (lebih gelap)
    # =========================================================================
    def _darken_color(self, color):
        if color.startswith('#'):
            color = color[1:]
        # Parse hex → int per channel
        r, g, b = int(color[:2], 16), int(color[2:4], 16), int(color[4:], 16)
        # Kurangi 30, pastikan tidak di bawah 0
        r, g, b = max(0, r-30), max(0, g-30), max(0, b-30)
        return f"#{r:02x}{g:02x}{b:02x}"

    # =========================================================================
    # METODE: show_about
    # Menampilkan halaman "About INVENTRA" di content area.
    # Kontennya adalah gambar inventraabt.png yang di-resize proporsional
    # mengikuti lebar content_container saat ini.
    # =========================================================================
    def show_about(self):
        # Bersihkan isi content area sebelumnya
        for widget in self.content_container.winfo_children():
            widget.destroy()

        wrapper = tk.Frame(self.content_container, bg=self.colors['bg'])
        wrapper.pack(fill="both", expand=True, padx=40, pady=40)

        try:
            img = Image.open(resource_path("inventraabt.png"))

            # Paksa Tkinter menghitung lebar aktual wrapper terlebih dahulu
            self.root.update_idletasks()
            available_width = wrapper.winfo_width()

            # Fallback jika winfo_width() belum bisa diukur (widget baru render)
            if available_width < 200:
                available_width = 1000

            # Resize proporsional (pertahankan aspect ratio)
            ratio      = available_width / img.width
            new_height = int(img.height * ratio)
            img        = img.resize((available_width, new_height), Image.LANCZOS)

            # Simpan referensi ke self agar tidak di-garbage collect Python
            # (ini adalah gotcha umum Tkinter — ImageTk harus disimpan di variabel instance)
            self.about_image = ImageTk.PhotoImage(img)

            tk.Label(
                wrapper,
                image=self.about_image,
                bg=self.colors['bg']
            ).pack()

        except Exception:
            # Jika gambar tidak ditemukan → tampilkan pesan teks
            tk.Label(
                wrapper,
                text="inventraabt.png tidak ditemukan.",
                font=("Segoe UI", 12),
                bg=self.colors['bg'],
                fg="red"
            ).pack()

        # Reset scroll ke atas setelah konten baru dimuat
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0)

    # =========================================================================
    # METODE: create_tutorial_screen
    # Membangun seluruh struktur UI tutorial:
    # Header → Sidebar (kiri) + Content Area (kanan, scrollable)
    # Dipanggil sekali saat __init__ dan hasilnya persisten selama tutorial terbuka.
    # =========================================================================
    def create_tutorial_screen(self):
        # Frame ini menimpa seluruh main window menggunakan .place()
        self.tutorial_frame = tk.Frame(self.root, bg=self.colors['bg'])
        self.tutorial_frame.place(x=0, y=0, relwidth=1, relheight=1)

        # ── HEADER ────────────────────────────────────────────────────────────
        header = tk.Frame(self.tutorial_frame, bg=self.colors['primary'], height=65)
        header.pack(fill="x")
        header.pack_propagate(False)  # paksa tinggi tetap 65px

        tk.Label(
            header,
            text="Quick Tutorial",
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['primary'],
            fg="white"
        ).pack(side="left", padx=30, pady=20)

        # ── Tombol Close dengan efek hover 3-state ────────────────────────────
        # 3 state: normal (primary) → hover (primary_dark) → klik (lebih gelap lagi)
        normal_color = self.colors['primary']
        hover_color  = self.colors['primary_dark']
        click_color  = self._darken_color(hover_color)

        close_btn = tk.Button(
            header,
            text="✕ Close",
            font=("Segoe UI", 12, "bold"),
            command=self.close_tutorial,
            bg=normal_color,
            fg="white",
            bd=0,
            activebackground=click_color,   # warna saat tombol ditekan (built-in Tkinter)
            activeforeground="white",
            relief="flat",
            cursor="hand2",
        )
        close_btn.pack(side="right", padx=30)

        # Bind manual untuk hover effect (Tkinter tidak punya built-in hover)
        close_btn.bind("<Enter>",          lambda e: close_btn.config(bg=hover_color))
        close_btn.bind("<Leave>",          lambda e: close_btn.config(bg=normal_color))
        close_btn.bind("<ButtonPress-1>",  lambda e: close_btn.config(bg=click_color))
        close_btn.bind("<ButtonRelease-1>",lambda e: close_btn.config(bg=hover_color))

        # ── MAIN AREA ─────────────────────────────────────────────────────────
        main_frame = tk.Frame(self.tutorial_frame, bg=self.colors['bg'])
        main_frame.pack(fill="both", expand=True)

        # ── SIDEBAR (KIRI) ────────────────────────────────────────────────────
        # Lebar fixed 260px, warna krem (#f5f1e8), border kanan abu-abu
        sidebar = tk.Frame(
            main_frame,
            bg="#f5f1e8",
            width=260,
            highlightbackground="#e5e7eb",
            highlightthickness=1
        )
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)  # paksa lebar tetap 260px

        # ── Tombol sidebar: About ─────────────────────────────────────────────
        # Setiap tombol sidebar menggunakan pola yang sama:
        # bg putih dengan hover abu-abu muda, anchor="w" agar teks rata kiri
        wrapper = tk.Frame(sidebar, bg="#f5f1e8")
        wrapper.pack(fill="x", padx=20, pady=6)
        indicator = tk.Frame(wrapper, width=4, bg="#dc2626")  # garis merah kiri (dekorasi)
        indicator.pack(side="left", fill="y")

        about_btn = tk.Button(
            sidebar,
            text="ⓘ  About Inventra",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff", fg="#1f2937",
            relief="flat", bd=0,
            highlightthickness=0,
            cursor="hand2",
            anchor="w", padx=20,
            command=self.show_about  # klik → tampilkan halaman About
        )
        about_btn.pack(fill="x", padx=20, pady=8, ipady=12)
        about_btn.bind("<Enter>", lambda e: about_btn.config(bg="#f3f4f6"))
        about_btn.bind("<Leave>", lambda e: about_btn.config(bg="#ffffff"))

        # ── Tombol sidebar: Tutorial ──────────────────────────────────────────
        wrapper = tk.Frame(sidebar, bg="#f5f1e8")
        wrapper.pack(fill="x", padx=20, pady=6)
        indicator = tk.Frame(wrapper, width=4, bg="#dc2626")
        indicator.pack(side="left", fill="y")

        tutorial_btn = tk.Button(
            sidebar,
            text="Tutorial    >",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff", fg="#1f2937",
            relief="flat", bd=0,
            highlightthickness=0,
            cursor="hand2",
            anchor="w", padx=20,
            command=self.render_all_steps  # klik → tampilkan semua langkah tutorial
        )
        tutorial_btn.pack(fill="x", padx=20, pady=8, ipady=12)
        tutorial_btn.bind("<Enter>", lambda e: tutorial_btn.config(bg="#f3f4f6"))
        tutorial_btn.bind("<Leave>", lambda e: tutorial_btn.config(bg="#ffffff"))

        # ── Tombol sidebar: Logika ────────────────────────────────────────────
        wrapper = tk.Frame(sidebar, bg="#f5f1e8")
        wrapper.pack(fill="x", padx=20, pady=6)
        indicator = tk.Frame(wrapper, width=4, bg="#dc2626")
        indicator.pack(side="left", fill="y")

        logika_btn = tk.Button(
            sidebar,
            text="Logika    >",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff", fg="#1f2937",
            relief="flat", bd=0,
            highlightthickness=0,
            cursor="hand2",
            anchor="w", padx=20,
            command=self.show_rumus  # klik → tampilkan slide rumus
        )
        logika_btn.pack(fill="x", padx=20, pady=8, ipady=12)
        logika_btn.bind("<Enter>", lambda e: logika_btn.config(bg="#f8fafc"))
        logika_btn.bind("<Leave>", lambda e: logika_btn.config(bg="#ffffff"))

        # ── CONTENT AREA (KANAN) dengan SCROLLABLE CANVAS ────────────────────
        # Pola scrollable Canvas di Tkinter:
        #   Canvas → create_window(scrollable_frame) → Scrollbar
        # Canvas adalah "viewport", scrollable_frame adalah konten sebenarnya.
        # Saat frame membesar, Canvas.bbox("all") diupdate sebagai scrollregion.
        content_area      = tk.Frame(main_frame, bg=self.colors['bg'])
        content_area.pack(side="right", fill="both", expand=True)

        container_wrapper = tk.Frame(content_area, bg=self.colors['bg'])
        container_wrapper.pack(fill="both", expand=True, padx=20, pady=20)

        self.canvas    = tk.Canvas(container_wrapper, bg="white", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(container_wrapper, orient="vertical", command=self.canvas.yview)

        # Frame yang benar-benar memuat widget (di dalam canvas)
        self.content_container = tk.Frame(self.canvas, bg="white")

        # Setiap kali content_container berubah ukuran → update scrollregion canvas
        self.content_container.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # Tempatkan content_container di pojok kiri atas canvas
        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.content_container,
            anchor="nw"
        )

        # Saat lebar canvas berubah → update lebar content_container agar mengikuti
        self.canvas.bind("<Configure>", self._resize_content)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Mouse wheel scroll — bind_all agar bisa scroll tanpa harus klik canvas dulu
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Default: tampilkan halaman About saat tutorial dibuka
        self.show_about()

    # =========================================================================
    # METODE: close_tutorial
    # Menutup tutorial dengan menghapus seluruh frame dari widget tree.
    # Karena frame ini di-place() di atas main window, main window langsung
    # kembali terlihat saat frame di-destroy.
    # =========================================================================
    def close_tutorial(self):
        self.tutorial_frame.destroy()

    # =========================================================================
    # METODE: render_all_steps
    # Menampilkan semua langkah tutorial dalam satu halaman scrollable.
    # Setiap langkah = satu "card" dengan title, deskripsi, gambar, dan tips.
    #
    # TEKNIK:
    #   1. Hapus semua widget lama di content_container
    #   2. Loop setiap step → buat card → pack ke content_container
    #   3. Reset scroll ke atas
    # =========================================================================
    def render_all_steps(self):
        # Hapus konten sebelumnya
        for widget in self.content_container.winfo_children():
            widget.destroy()

        for step in self.tutorial_steps:
            # ── Card container per step ───────────────────────────────────────
            # Setiap card diberi border abu-abu tipis untuk kesan "card"
            card = tk.Frame(
                self.content_container,
                bg="white",
                highlightbackground="#e5e7eb",
                highlightthickness=1
            )
            card.pack(fill="x", padx=100, pady=30)

            # Padding dalam card
            section = tk.Frame(card, bg="white")
            section.pack(fill="both", expand=True, padx=30, pady=25)

            # ── Judul step ────────────────────────────────────────────────────
            tk.Label(
                section,
                text=step["title"],
                font=("Segoe UI", 18, "bold"),
                bg="white",
                fg=self.colors['text']
            ).pack(anchor="w", pady=(0, 12))

            # ── Deskripsi ─────────────────────────────────────────────────────
            # wraplength=850 agar teks tidak keluar dari card
            tk.Label(
                section,
                text=step["description"],
                font=("Segoe UI", 12),
                bg="white",
                fg=self.colors['text'],
                justify="left",
                wraplength=850
            ).pack(anchor="w", pady=(0, 20))

            # ── Gambar ────────────────────────────────────────────────────────
            image_frame = tk.Frame(section, bg="white")
            image_frame.pack(fill="x", pady=(0, 20))

            try:
                img = Image.open(resource_path(step["image"]))
                # thumbnail() = resize maksimal 950x450 sambil jaga aspect ratio
                img.thumbnail((950, 450), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)

                img_label = tk.Label(image_frame, image=photo, bg="white")
                # PENTING: simpan referensi photo di widget agar tidak di-GC
                img_label.image = photo
                img_label.pack()

            except Exception:
                # Gambar tidak ditemukan → tampilkan pesan placeholder
                tk.Label(
                    image_frame,
                    text=f"Gambar tidak ditemukan:\n{step['image']}",
                    font=("Segoe UI", 11),
                    bg="white",
                    fg=self.colors['text_light']
                ).pack(pady=20)

            # ── Kotak Tips ────────────────────────────────────────────────────
            # Latar belakang primary_light (warna muda dari tema merah)
            tips_frame = tk.Frame(section, bg=self.colors['primary_light'])
            tips_frame.pack(fill="x", pady=(10, 0))

            tk.Label(
                tips_frame,
                text=step["tips"],
                font=("Segoe UI", 11),
                bg=self.colors['primary_light'],
                fg=self.colors['text'],
                anchor="w"
            ).pack(fill="x", padx=15, pady=12)

        # Reset scroll ke atas setelah semua card dirender
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0)

    # =========================================================================
    # METODE: _resize_content
    # Event handler yang dipanggil saat lebar Canvas berubah (misal: window di-resize).
    # Memastikan content_container mengikuti lebar canvas, bukan menjadi lebih sempit.
    # =========================================================================
    def _resize_content(self, event):
        # Update lebar window di dalam canvas agar = lebar canvas saat ini
        self.canvas.itemconfig(self.canvas_window, width=event.width)

    # =========================================================================
    # METODE: _on_mousewheel
    # Handler scroll mouse. event.delta = 120 per "klik" scroll di Windows.
    # Dibagi 120 → 1 unit scroll. Dikali -1 karena scroll atas = delta positif
    # tapi canvas harus bergerak ke atas (yview negatif).
    # =========================================================================
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    # =========================================================================
    # METODE: show_rumus
    # Menampilkan slide-slide PNG dari folder rumus_slide/ secara berurutan.
    # Slide ini menjelaskan logika matematis perhitungan ROP/ROQ/KLASIFIKASI.
    #
    # Folder rumus_slide/ bisa diisi dengan PNG sebanyak apapun.
    # File diurutkan secara alfabetis (sorted()) sehingga urutan penamaan file
    # menentukan urutan tampil (misal: 01_intro.png, 02_rumus.png, dll)
    # =========================================================================
    def show_rumus(self):
        # Bersihkan konten sebelumnya
        for widget in self.content_container.winfo_children():
            widget.destroy()

        content = tk.Frame(self.content_container, bg="white")
        content.pack(fill="both", expand=True, padx=60, pady=40)

        tk.Label(
            content,
            text="LOGIKA PERHITUNGAN INVENTORY",
            font=("Segoe UI", 20, "bold"),
            bg="white",
            fg=self.colors['text']
        ).pack(anchor="w", pady=(0, 30))

        folder_path = resource_path("rumus_slide")

        if not os.path.exists(folder_path):
            tk.Label(
                content,
                text="Folder rumus_slide tidak ditemukan.",
                font=("Segoe UI", 12),
                bg="white",
                fg="red"
            ).pack()
            return

        # List untuk menyimpan referensi semua ImageTk.PhotoImage
        # Jika tidak disimpan di self, Python GC akan menghapusnya dan gambar hilang
        self.rumus_images = []

        # Ambil hanya file .png, urutkan alfabetis
        slides = sorted([f for f in os.listdir(folder_path) if f.endswith(".png")])

        for slide in slides:
            img_path = os.path.join(folder_path, slide)
            img      = Image.open(img_path)
            # thumbnail dengan batas besar (1000x2000) → hanya resize jika terlalu besar
            img.thumbnail((1000, 2000), Image.LANCZOS)

            photo = ImageTk.PhotoImage(img)
            self.rumus_images.append(photo)  # simpan referensi agar tidak di-GC

            tk.Label(
                content,
                image=photo,
                bg="white"
            ).pack(pady=20)

        # Reset scroll ke atas
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0)