import tkinter as tk
from PIL import Image, ImageTk
import sys
import os


def resource_path(relative_path):
    """Return path yang benar saat jadi EXE atau dijalankan di Python biasa"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class TutorialScreen:
    def __init__(self, root, colors, on_continue):
        self.root = root
        self.colors = colors
        self.on_continue = on_continue
        self.current_step = 0
        
        # Tutorial steps dengan gambar
        self.tutorial_steps = [
            {
                "title": "1. Upload File Excel",
                "description": "Klik tombol '+ Upload File' dan pilih file Excel (.xlsx/.xls)\nAnda bisa pilih beberapa file sekaligus",
                "image": "tutorial_upload.png",  # ganti dengan nama file gambar Anda
                "tips": "💡 Pastikan file Excel tidak sedang dibuka di aplikasi lain"
            },
            {
                "title": "2. Process & Mapping Data",
                "description": "Klik '⚙ Process', lalu pilih Source File dan Sheet Name untuk setiap tab\nAtur nama kolom sesuai dengan Excel Anda",
                "image": "tutorial_mapping.png",
                "tips": "⚠️ Nama kolom harus PERSIS sama dengan di Excel!"
            },
            {
                "title": "3. Preview & Save",
                "description": "Lihat hasil di Data Preview, lalu klik '💾 Save Excel' untuk menyimpan",
                "image": "tutorial_preview.png",
                "tips": "✅ Selalu periksa preview sebelum menyimpan"
            },
            {
                "title": "4. Reset Data",
                "description": "Klik tombol 'Reset All' untuk membersihkan semua file dan hasil preview.\nGunakan ini jika ingin memulai ulang dari awal.",
                "image": "tutorial_reset.png",  # siapkan gambar ini
                "tips": "🗑️ Reset Data akan menghapus semua file yang sudah diupload dan hasil proses."
            }  
           ]  
        
        self.create_tutorial_screen()

    def _darken_color(self, color):
        """Darken a hex color"""
        if color.startswith('#'):
            color = color[1:]
        r, g, b = int(color[:2], 16), int(color[2:4], 16), int(color[4:], 16)
        r, g, b = max(0, r-30), max(0, g-30), max(0, b-30)
        return f"#{r:02x}{g:02x}{b:02x}"

    def show_about(self):
        for widget in self.content_container.winfo_children():
            widget.destroy()

        wrapper = tk.Frame(self.content_container, bg=self.colors['bg'])
        wrapper.pack(fill="both", expand=True, padx=40, pady=40)

        try:
            img = Image.open(resource_path("inventraabt.png"))

            # 🔥 Pakai lebar wrapper, bukan canvas
            self.root.update_idletasks()
            available_width = wrapper.winfo_width()

            if available_width < 200:
                available_width = 1000  # fallback aman

            # Resize proporsional
            ratio = available_width / img.width
            new_height = int(img.height * ratio)

            img = img.resize((available_width, new_height), Image.LANCZOS)

            self.about_image = ImageTk.PhotoImage(img)

            tk.Label(
                wrapper,
                image=self.about_image,
                bg=self.colors['bg']
            ).pack()

        except Exception:
            tk.Label(
                wrapper,
                text="inventraabt.png tidak ditemukan.",
                font=("Segoe UI", 12),
                bg=self.colors['bg'],
                fg="red"
            ).pack()

        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0)


    
    def create_tutorial_screen(self):
        self.tutorial_frame = tk.Frame(self.root, bg=self.colors['bg'])
        self.tutorial_frame.place(x=0, y=0, relwidth=1, relheight=1)

        # ===== HEADER =====
        header = tk.Frame(self.tutorial_frame, bg=self.colors['primary'], height=65)
        header.pack(fill="x")
        header.pack_propagate(False)
        
        tk.Label(
            header,
            text="Quick Tutorial",
            font=("Segoe UI", 16, "bold"),
            bg=self.colors['primary'],
            fg="white"
        ).pack(side="left", padx=30, pady=20)

        # === EXIT BUTTON MODERN STATE ===
        normal_color = self.colors['primary']               # warna header
        hover_color = self.colors['primary_dark']          # saat hover
        click_color = self._darken_color(hover_color)      # lebih gelap lagi

        close_btn = tk.Button(
            header,
            text="✕ Close",
            font=("Segoe UI", 12, "bold"),
            command=self.close_tutorial,
            bg=normal_color,
            fg="white",
            bd=0,
            activebackground=click_color,   # supaya tidak jadi putih saat klik
            activeforeground="white",
            relief="flat",
            cursor="hand2",
        )

        close_btn.pack(side="right", padx=30)

        # Hover
        close_btn.bind("<Enter>", lambda e: close_btn.config(bg=hover_color))

        # Leave
        close_btn.bind("<Leave>", lambda e: close_btn.config(bg=normal_color))

        # Saat ditekan
        close_btn.bind("<ButtonPress-1>", lambda e: close_btn.config(bg=click_color))

        # Saat dilepas
        close_btn.bind("<ButtonRelease-1>", lambda e: close_btn.config(bg=hover_color))

        # Hover effect
        # close_btn.bind("<Enter>", lambda e: close_btn.config(bg="#374151"))
        # close_btn.bind("<Leave>", lambda e: close_btn.config(bg="#1f2937"))






        # tk.Button(
        #     header,
        #     text="← Back",
        #     font=("Segoe UI", 11, "bold"),
        #     command=self.skip_tutorial,
        #     bg=self.colors['primary'],
        #     fg="white",
        #     relief="flat",
        #     cursor="hand2",
        #     activebackground=self.colors['primary'],
        #     activeforeground="white",
        # ).pack(side="right", padx=30)


        # ===== MAIN AREA =====
        main_frame = tk.Frame(self.tutorial_frame, bg=self.colors['bg'])
        main_frame.pack(fill="both", expand=True)

        # ===== SIDEBAR (KIRI) =====
        sidebar = tk.Frame(
            main_frame,
            bg="#f5f1e8",
            width=260,
            highlightbackground="#e5e7eb",
            highlightthickness=1
        )


        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        wrapper = tk.Frame(sidebar, bg="#f5f1e8")
        wrapper.pack(fill="x", padx=20, pady=6)

        indicator = tk.Frame(wrapper, width=4, bg="#dc2626")
        indicator.pack(side="left", fill="y")

        about_btn = tk.Button(
            sidebar,
            text="ⓘ  About Inventra",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff",
            fg="#1f2937",
            relief="flat",
            bd=0,
            highlightthickness=0,
            cursor="hand2",
             anchor="w",
             padx=20,
            command=self.show_about
        )

        about_btn.pack(fill="x", padx=20, pady=8, ipady=12)

        # Hover effect
        about_btn.bind("<Enter>", lambda e: about_btn.config(bg="#f3f4f6"))
        about_btn.bind("<Leave>", lambda e: about_btn.config(bg="#ffffff"))

        wrapper = tk.Frame(sidebar, bg="#f5f1e8")
        wrapper.pack(fill="x", padx=20, pady=6)

        indicator = tk.Frame(wrapper, width=4, bg="#dc2626")
        indicator.pack(side="left", fill="y")

        tutorial_btn = tk.Button(
            sidebar,
            text="Tutorial    >",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff",
            fg="#1f2937",
            relief="flat",
            bd=0,
            highlightthickness=0,
            cursor="hand2",
            anchor="w",
            padx=20,
            command=self.render_all_steps
        )

        tutorial_btn.pack(fill="x", padx=20, pady=8, ipady=12)

        tutorial_btn.bind("<Enter>", lambda e: tutorial_btn.config(bg="#f3f4f6"))
        tutorial_btn.bind("<Leave>", lambda e: tutorial_btn.config(bg="#ffffff"))

        wrapper = tk.Frame(sidebar, bg="#f5f1e8")
        wrapper.pack(fill="x", padx=20, pady=6)

        indicator = tk.Frame(wrapper, width=4, bg="#dc2626")
        indicator.pack(side="left", fill="y")

        logika_btn = tk.Button(
            sidebar,
            text="Logika    >",
            font=("Segoe UI", 11, "bold"),
            bg="#ffffff",
            fg="#1f2937",
            relief="flat",
            bd=0,
            highlightthickness=0,
            cursor="hand2",
             anchor="w",
            padx=20,
            command=self.show_rumus
        )
        logika_btn.pack(fill="x", padx=20, pady=8, ipady=12)

        logika_btn.bind("<Enter>", lambda e: logika_btn.config(bg="#f8fafc"))
        logika_btn.bind("<Leave>", lambda e: logika_btn.config(bg="#ffffff"))




        # ===== CONTENT AREA (KANAN) =====
        content_area = tk.Frame(main_frame, bg=self.colors['bg'])
        content_area.pack(side="right", fill="both", expand=True)

        container_wrapper = tk.Frame(content_area, bg=self.colors['bg'])

        container_wrapper.pack(fill="both", expand=True, padx=20, pady=20)

        self.canvas = tk.Canvas(container_wrapper, bg="white", highlightthickness=0)
        self.scrollbar = tk.Scrollbar(container_wrapper, orient="vertical", command=self.canvas.yview)

        self.content_container = tk.Frame(self.canvas, bg="white")

        self.content_container.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas_window = self.canvas.create_window(
            (0, 0),
            window=self.content_container,
            anchor="nw"
        )

        self.canvas.bind("<Configure>", self._resize_content)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Enter>", lambda e: self.canvas.bind_all("<MouseWheel>", self._on_mousewheel))
        self.canvas.bind("<Leave>", lambda e: self.canvas.unbind_all("<MouseWheel>"))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)


        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

        # Default tampil tutorial
        self.show_about()

    def close_tutorial(self):
        self.tutorial_frame.destroy()

    
    def render_all_steps(self):

        # Bersihkan isi lama
        for widget in self.content_container.winfo_children():
            widget.destroy()

        for step in self.tutorial_steps:

            # ===== CARD CONTAINER =====
            card = tk.Frame(
                self.content_container,
                bg="white",
                highlightbackground="#e5e7eb",
                highlightthickness=1
            )
            card.pack(fill="x", padx=100, pady=30)

            # Inner spacing
            section = tk.Frame(card, bg="white")
            section.pack(fill="both", expand=True, padx=30, pady=25)

            # ===== TITLE =====
            tk.Label(
                section,
                text=step["title"],
                font=("Segoe UI", 18, "bold"),
                bg="white",
                fg=self.colors['text']
            ).pack(anchor="w", pady=(0, 12))

            # ===== DESCRIPTION =====
            tk.Label(
                section,
                text=step["description"],
                font=("Segoe UI", 12),
                bg="white",
                fg=self.colors['text'],
                justify="left",
                wraplength=850
            ).pack(anchor="w", pady=(0, 20))

            # ===== IMAGE =====
            image_frame = tk.Frame(section, bg="white")
            image_frame.pack(fill="x", pady=(0, 20))

            try:
                img = Image.open(resource_path(step["image"]))
                img.thumbnail((950, 450), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)

                img_label = tk.Label(image_frame, image=photo, bg="white")
                img_label.image = photo
                img_label.pack()

            except Exception:
                tk.Label(
                    image_frame,
                    text=f"Gambar tidak ditemukan:\n{step['image']}",
                    font=("Segoe UI", 11),
                    bg="white",
                    fg=self.colors['text_light']
                ).pack(pady=20)

            # ===== TIPS BOX =====
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

        # Reset scroll
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0)


    def _resize_content(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)


    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def show_rumus(self):
        # Bersihkan area
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

        self.rumus_images = []

        slides = sorted([f for f in os.listdir(folder_path) if f.endswith(".png")])

        for slide in slides:
            img_path = os.path.join(folder_path, slide)
            img = Image.open(img_path)
            img.thumbnail((1000, 2000), Image.LANCZOS)

            photo = ImageTk.PhotoImage(img)
            self.rumus_images.append(photo)

            tk.Label(
                content,
                image=photo,
                bg="white"
            ).pack(pady=20)

        # 🔥 Reset scroll ke atas
        self.canvas.update_idletasks()
        self.canvas.yview_moveto(0)

    # def skip_tutorial(self):
    #     self.finish_tutorial()

    # def finish_tutorial(self):
    #     self.tutorial_frame.destroy()
    #     self.on_continue()


