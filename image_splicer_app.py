import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk, ImageGrab, ImageDraw, ImageFont # Added ImageDraw, ImageFont
import os # 用于检查文件路径
import tempfile # Added tempfile
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:
    TkinterDnD = None # Fallback if not installed
    DND_FILES = None
    messagebox.showwarning("警告", "tkinterdnd2 模块未找到。\n拖拽功能将不可用。\n请运行: pip install tkinterdnd2")


class ImageSplicerApp:
    def __init__(self, root_tk_or_dnd): # root can be tk.Tk or TkinterDnD.Tk
        self.root = root_tk_or_dnd
        self.root.title("图片拼接工具")
        self.root.geometry("800x600")

        self.image_paths = []
        self.processed_image = None # 用于存储拼接后的Pillow Image对象
        self.processed_image_tk = None # 用于在Tkinter中显示的ImageTk.PhotoImage对象

        # --- 左侧：图片列表和控制 ---
        left_frame = ttk.Frame(root, padding="10")
        left_frame.grid(row=0, column=0, sticky="nswe")
        root.grid_columnconfigure(0, weight=1)
        root.grid_rowconfigure(0, weight=1)

        # 图片列表
        ttk.Label(left_frame, text="待拼接图片:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.listbox_images = tk.Listbox(left_frame, selectmode=tk.EXTENDED, width=40, height=15)
        self.listbox_images.grid(row=1, column=0, columnspan=2, sticky="nsew")
        left_frame.grid_rowconfigure(1, weight=1)

        # 注册拖拽目标
        if TkinterDnD and DND_FILES:
            self.listbox_images.drop_target_register(DND_FILES)
            self.listbox_images.dnd_bind('<<Drop>>', self.handle_drop)

        # 绑定粘贴事件
        self.root.bind_all('<Control-v>', self.handle_paste) # Bind to root for global paste
        # Or bind specifically to listbox if preferred:
        # self.listbox_images.bind('<Control-v>', self.handle_paste)


        # 控制按钮
        btn_frame = ttk.Frame(left_frame)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=(10,0), sticky="ew")

        self.btn_add = ttk.Button(btn_frame, text="添加图片", command=self.add_images)
        self.btn_add.pack(side=tk.LEFT, padx=5)

        self.btn_remove = ttk.Button(btn_frame, text="移除选中", command=self.remove_selected_images)
        self.btn_remove.pack(side=tk.LEFT, padx=5)

        self.btn_clear = ttk.Button(btn_frame, text="清空列表", command=self.clear_images)
        self.btn_clear.pack(side=tk.LEFT, padx=5)

        # 拼接模式
        ttk.Label(left_frame, text="拼接模式:").grid(row=3, column=0, sticky="w", pady=(10, 5))
        self.splice_mode_var = tk.StringVar(value="横向拼接")
        self.combo_splice_mode = ttk.Combobox(left_frame, textvariable=self.splice_mode_var,
                                              values=["横向拼接", "纵向拼接", "2xN网格"], state="readonly")
        self.combo_splice_mode.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(0,10))

        # 水印开关和颜色选择
        watermark_frame = ttk.Frame(left_frame)
        watermark_frame.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(5,0))
        
        self.watermark_var = tk.BooleanVar(value=False)
        self.chk_watermark = ttk.Checkbutton(watermark_frame, text="添加序号水印", variable=self.watermark_var)
        self.chk_watermark.pack(side=tk.LEFT)

        self.watermark_color_var = tk.StringVar(value="红色") # Default color
        self.combo_watermark_color = ttk.Combobox(watermark_frame, textvariable=self.watermark_color_var,
                                                 values=["红色", "白色", "黑色"], state="readonly", width=5)
        self.combo_watermark_color.pack(side=tk.LEFT, padx=(5,0))
        ttk.Label(watermark_frame, text="(左上角)").pack(side=tk.LEFT, padx=(2,0))


        # 开始拼接按钮
        self.btn_splice = ttk.Button(left_frame, text="开始拼接", command=self.splice_images)
        self.btn_splice.grid(row=6, column=0, columnspan=2, pady=10, sticky="ew")

        # 左侧选中图片缩略图预览
        ttk.Label(left_frame, text="选中项预览:").grid(row=7, column=0, columnspan=2, sticky="w", pady=(10, 5))
        self.thumbnail_canvas = tk.Canvas(left_frame, bg="lightgrey", width=150, height=150) # 缩略图区域大小
        self.thumbnail_canvas.grid(row=8, column=0, columnspan=2, pady=(0,10), sticky="ew")
        self.thumbnail_tk = None # To hold the PhotoImage for the thumbnail

        # 绑定列表框选择事件
        self.listbox_images.bind('<<ListboxSelect>>', self.update_thumbnail_preview)


        # --- 右侧：预览和输出 ---
        right_frame = ttk.Frame(root, padding="10")
        right_frame.grid(row=0, column=1, sticky="nswe")
        root.grid_columnconfigure(1, weight=3) # 让右侧区域更宽

        ttk.Label(right_frame, text="预览:").grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.preview_canvas = tk.Canvas(right_frame, bg="lightgrey", width=400, height=400) # 初始大小
        self.preview_canvas.grid(row=1, column=0, sticky="nsew")
        right_frame.grid_rowconfigure(1, weight=1)
        right_frame.grid_columnconfigure(0, weight=1)


        # 输出按钮
        output_btn_frame = ttk.Frame(right_frame)
        output_btn_frame.grid(row=2, column=0, pady=(10,0), sticky="ew")

        self.btn_copy = ttk.Button(output_btn_frame, text="复制图片", command=self.copy_image, state=tk.DISABLED)
        self.btn_copy.pack(side=tk.LEFT, padx=5)
        
        # 保存格式选择
        self.save_format_var = tk.StringVar(value="PNG") # Default to PNG
        self.combo_save_format = ttk.Combobox(output_btn_frame, textvariable=self.save_format_var,
                                             values=["PNG", "JPG"], state="readonly", width=5)
        self.combo_save_format.pack(side=tk.LEFT, padx=(5,0))


        self.btn_save = ttk.Button(output_btn_frame, text="保存图片", command=self.save_image, state=tk.DISABLED)
        self.btn_save.pack(side=tk.LEFT, padx=5)


    def _add_single_image_path(self, file_path):
        """Helper function to add a single valid image path to the list."""
        # Basic validation for file extension (can be expanded)
        valid_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp')
        if not os.path.isfile(file_path) or not file_path.lower().endswith(valid_extensions):
            # Optionally show a warning for invalid files, or silently ignore
            # messagebox.showwarning("无效文件", f"文件 '{file_path.split('/')[-1]}' 不是支持的图片格式或不存在。")
            return False

        if file_path not in self.image_paths:
            self.image_paths.append(file_path)
            self.listbox_images.insert(tk.END, os.path.basename(file_path))
            return True
        return False

    def add_images(self):
        files = filedialog.askopenfilenames(
            title="选择图片",
            filetypes=(("图片文件", "*.jpg *.jpeg *.png *.bmp *.gif *.webp"), ("所有文件", "*.*"))
        )
        if files:
            for f_path in files:
                self._add_single_image_path(f_path)

    def handle_drop(self, event):
        # event.data is a string containing one or more file paths
        # Paths might be space-separated, and if multiple, often enclosed in {}
        raw_paths = event.data
        if not raw_paths:
            return

        # Clean up the string: remove curly braces if present
        if raw_paths.startswith('{') and raw_paths.endswith('}'):
            raw_paths = raw_paths[1:-1]

        # Split paths. This might need adjustment based on how TkinterDnD2 formats multiple files.
        # A common way is space-separated, but paths with spaces can be tricky.
        # TkinterDnD2 often gives one path per drop event if multiple files are selected and dropped together,
        # or a tcl list like string. Let's try a robust split.
        
        # This is a common way TkinterDnD2 provides paths (as a Tcl-formatted list string)
        try:
            # Attempt to parse as a Tcl list string if it contains spaces and isn't a single simple path
            if ' ' in raw_paths and not os.path.exists(raw_paths): # if it's not a single path with spaces
                 # Use Tk's built-in Tcl interpreter to split the list properly
                paths_to_add = self.root.tk.splitlist(raw_paths)
            else: # Assume single path or paths without spaces that don't need Tcl parsing
                paths_to_add = [raw_paths] if os.path.exists(raw_paths) else raw_paths.split()


        except tk.TclError: # Fallback for simple space split if Tcl parsing fails
            paths_to_add = raw_paths.split()


        added_count = 0
        for p in paths_to_add:
            # Sometimes paths from drag-drop might have extra quotes, strip them
            p_cleaned = p.strip('"\'')
            if self._add_single_image_path(p_cleaned):
                added_count +=1
        
        if added_count == 0 and paths_to_add: # If paths were provided but none were valid/added
            messagebox.showwarning("拖拽提示", "未能添加任何拖拽的图片。\n请确保文件是支持的图片格式。")


    def handle_paste(self, event=None): # event is passed by binding but not always used
        added_count = 0
        # 1. Try to paste as file paths first
        try:
            clipboard_content = self.root.clipboard_get()
            potential_paths = clipboard_content.replace('\r\n', '\n').split('\n')
            for p_raw in potential_paths:
                p_trimmed = p_raw.strip()
                if p_trimmed:
                    p_cleaned = p_trimmed.strip('"\'')
                    if self._add_single_image_path(p_cleaned):
                        added_count += 1
        except tk.TclError:
            # This error means clipboard doesn't contain text,
            # so we can proceed to try grabbing an image.
            pass # Silently ignore, will try ImageGrab next.
        except Exception as e:
            # Other unexpected error while processing text
            # messagebox.showerror("粘贴错误", f"处理剪贴板文本时出错: {e}")
            pass


        # 2. If no paths were added, try to paste as image data
        if added_count == 0:
            try:
                im = ImageGrab.grabclipboard()
                if im:
                    # Create a temporary file to save the image
                    # Suffix is important for Pillow to know the format.
                    # Delete=False is important on Windows, otherwise the file might be
                    # locked or deleted before Image.open() can access it in _add_single_image_path.
                    # We will manage these temp files if needed, or let OS handle them.
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".png", prefix="pasted_img_") as tmp_file:
                        # Ensure image is in a saveable mode (e.g. RGB for PNG if it's RGBA)
                        if im.mode == 'RGBA' or im.mode == 'P':
                             # PNG can handle RGBA, but let's convert to RGB if P for wider compatibility
                             # or if we want to ensure no alpha issues if not handled later.
                             # For now, let's save as is if PNG.
                             pass
                        elif im.mode != 'RGB' and im.mode != 'L': # L for grayscale
                            im = im.convert('RGB')
                        
                        im.save(tmp_file.name)
                        temp_file_path = tmp_file.name
                    
                    # Now add this temporary file's path
                    if self._add_single_image_path(temp_file_path):
                        # Optionally, store temp_file_path in a list to clean up later if desired
                        # e.g., self.temporary_files.append(temp_file_path)
                        # And rename it in the listbox to something like "Pasted Image 1"
                        # For now, it will show the temp file name.
                        # Let's try to give it a more user-friendly name in the listbox
                        base_name = os.path.basename(temp_file_path)
                        listbox_index = self.listbox_images.get(0, tk.END).index(base_name)
                        self.listbox_images.delete(listbox_index)
                        
                        # Find a unique name like "Pasted Image 1", "Pasted Image 2"
                        paste_idx = 1
                        while f"Pasted Image {paste_idx}.png" in [os.path.basename(p) for p in self.image_paths if "pasted_img_" not in p]: # Avoid clashing with actual filenames
                            # Check against existing non-temp pasted images
                            is_name_taken = False
                            for item_idx in range(self.listbox_images.size()):
                                if self.listbox_images.get(item_idx) == f"Pasted Image {paste_idx}.png":
                                    is_name_taken = True
                                    break
                            if is_name_taken:
                                paste_idx += 1
                            else:
                                break
                        
                        new_display_name = f"Pasted Image {paste_idx}.png"
                        self.listbox_images.insert(listbox_index, new_display_name)
                        
                        # Update the actual path in self.image_paths to reflect the new name if we were to rename the temp file
                        # For now, self.image_paths still holds the temp file path, listbox shows friendly name.
                        # This is a bit of a hack for display. A better way would be to store a tuple (display_name, actual_path)
                        # Or, rename the actual temp file (but that's more complex with tempfile).
                        # For simplicity, we'll just show a friendly name but the path remains temp.
                        
                        added_count += 1
                    else:
                        # If _add_single_image_path failed (e.g. somehow not a valid image after saving)
                        os.unlink(temp_file_path) # Clean up the temp file
                        messagebox.showwarning("粘贴提示", "无法处理粘贴的图片数据。")

                elif not clipboard_content: # If clipboard_get also failed and grabclipboard is None
                    messagebox.showwarning("粘贴提示", "剪贴板为空或不包含可识别的图片数据或文件路径。")

            except ImportError:
                 messagebox.showerror("粘贴错误", "Pillow (PIL) 的 ImageGrab 模块无法导入。\n请确保 Pillow 已正确安装。")
            except Exception as e:
                # This might catch errors if clipboard is not image data, e.g. on Linux without xclip/xsel
                # Or other ImageGrab errors.
                # Check if clipboard_content was empty from the first try
                if not clipboard_content: # From the first try_except block
                    messagebox.showwarning("粘贴提示", "剪贴板为空或不包含可识别的图片数据或文件路径。")
                # else:
                # An error occurred trying to grab image data, but there might have been text.
                # If added_count is still 0, it means neither path nor image worked.
                # The original message for path failure would have been suppressed.
                # So, if paths were present but invalid, and image grab failed, show a generic message.
                if added_count == 0 and clipboard_content and any(p.strip() for p in potential_paths) :
                     messagebox.showwarning("粘贴提示", "未能从剪贴板添加任何图片。\n请确保复制的是有效的图片文件路径或图片数据。")
                elif added_count == 0: # General fallback if clipboard was truly empty or unhandled format
                     messagebox.showwarning("粘贴提示", "无法从剪贴板获取图片数据。")
                # print(f"Error grabbing image from clipboard: {e}") # For debugging

        # Final check if anything was added
        if added_count == 0 and not (clipboard_content and any(p.strip() for p in potential_paths)) and not im :
             # This condition is a bit redundant due to messages above, but as a final catch-all
             # if clipboard_content was None/empty and im was None.
             pass # Messages should have been shown already.


    def remove_selected_images(self):
        selected_indices = self.listbox_images.curselection()
        if not selected_indices:
            messagebox.showwarning("提示", "请先选择要移除的图片。")
            return

        for i in sorted(selected_indices, reverse=True):
            self.listbox_images.delete(i)
            del self.image_paths[i]

    def clear_images(self):
        self.listbox_images.delete(0, tk.END)
        self.image_paths.clear()
        self.processed_image = None
        self.processed_image_tk = None
        self.preview_canvas.delete("all")
        self.thumbnail_canvas.delete("all") # 清空缩略图
        self.thumbnail_tk = None
        self.btn_copy.config(state=tk.DISABLED)
        self.btn_save.config(state=tk.DISABLED)

    def update_thumbnail_preview(self, event=None):
        self.thumbnail_canvas.delete("all")
        self.thumbnail_tk = None

        selected_indices = self.listbox_images.curselection()
        if not selected_indices:
            return

        try:
            selected_index = selected_indices[0] # Get the first selected item
            image_path = self.image_paths[selected_index]

            img = Image.open(image_path)
            
            canvas_w = self.thumbnail_canvas.winfo_width()
            canvas_h = self.thumbnail_canvas.winfo_height()
            if canvas_w <= 1: canvas_w = self.thumbnail_canvas.cget("width")
            if canvas_h <= 1: canvas_h = self.thumbnail_canvas.cget("height")
            if isinstance(canvas_w, str): canvas_w = int(canvas_w)
            if isinstance(canvas_h, str): canvas_h = int(canvas_h)


            img_w, img_h = img.size
            if img_w == 0 or img_h == 0: return

            ratio = min(canvas_w / img_w, canvas_h / img_h)
            
            new_w = int(img_w * ratio)
            new_h = int(img_h * ratio)

            if new_w > 0 and new_h > 0:
                img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                self.thumbnail_tk = ImageTk.PhotoImage(img_resized)
                
                # Center the image on the canvas
                x_pos = (canvas_w - new_w) / 2
                y_pos = (canvas_h - new_h) / 2
                self.thumbnail_canvas.create_image(x_pos, y_pos, anchor=tk.NW, image=self.thumbnail_tk)
            img.close()

        except Exception as e:
            # print(f"Error updating thumbnail: {e}") # For debugging
            # Could show a placeholder or error message on the thumbnail canvas
            pass


    def update_preview(self, pil_image):
        self.processed_image = pil_image
        self.preview_canvas.delete("all")
        
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()

        if canvas_width <= 1 or canvas_height <= 1:
            canvas_width = self.preview_canvas.cget("width") # Fallback to configured width
            canvas_height = self.preview_canvas.cget("height") # Fallback to configured height
            if isinstance(canvas_width, str): canvas_width = int(canvas_width)
            if isinstance(canvas_height, str): canvas_height = int(canvas_height)


        img_copy = pil_image.copy()
        img_width, img_height = img_copy.size
        
        if img_width == 0 or img_height == 0: return # Avoid division by zero

        ratio = min(canvas_width / img_width, canvas_height / img_height)
        
        if ratio < 1:
            new_width = int(img_width * ratio)
            new_height = int(img_height * ratio)
            if new_width > 0 and new_height > 0:
                 img_copy = img_copy.resize((new_width, new_height), Image.Resampling.LANCZOS)

        self.processed_image_tk = ImageTk.PhotoImage(img_copy)
        
        x = (canvas_width - img_copy.width) / 2
        y = (canvas_height - img_copy.height) / 2
        self.preview_canvas.create_image(x, y, anchor=tk.NW, image=self.processed_image_tk)
        
        self.btn_copy.config(state=tk.NORMAL)
        self.btn_save.config(state=tk.NORMAL)

    def _prepare_image_for_paste(self, image_object):
        """Converts image to RGBA to handle transparency correctly for an RGBA canvas."""
        if image_object.mode == 'RGBA':
            return image_object.copy() # Use a copy
        elif image_object.mode == 'P': # Palette mode
            # Convert to RGBA. This preserves transparency if defined in palette.
            return image_object.convert('RGBA')
        else: # RGB, L, CMYK etc.
            # Convert to RGBA. Opaque pixels get alpha=255.
            return image_object.convert('RGBA')

    def splice_images(self):
        if not self.image_paths:
            messagebox.showwarning("提示", "请先添加图片。")
            return

        images_to_splice_orig = [] # To store original Pillow image objects for closing
        images_to_splice = []      # To store RGBA prepared Pillow image objects

        try:
            for p in self.image_paths:
                img_orig = Image.open(p)
                images_to_splice_orig.append(img_orig) # Keep original for closing
                images_to_splice.append(self._prepare_image_for_paste(img_orig))

        except Exception as e:
            messagebox.showerror("错误", f"打开或预处理图片失败: {e}")
            # Close any successfully opened images if error occurs mid-list
            for img_obj in images_to_splice_orig: # Close the opened original images
                img_obj.close()
            return
        finally:
            # Always close original Pillow image objects after they've been processed or if an error occurs
            for img_obj in images_to_splice_orig:
                img_obj.close()
        
        # Add watermark if enabled
        if self.watermark_var.get():
            images_with_watermark = []
            font = None
            font_size = 20
            try:
                # Try to load a common font, adjust path if necessary or use default
                # For simplicity, let's try a very basic font loading.
                # A more robust solution would bundle a font or use a font finding mechanism.
                font_path_msyh = "msyh.ttf" # Assuming "微软雅黑.ttf" is named this and in the same directory
                font_path_simsun = "simsun.ttc" #宋体, common on Windows
                font_path_arial = "arial.ttf"
                font_path_dejavu = "DejaVuSans.ttf"

                font_load_attempts = [font_path_msyh, font_path_simsun, font_path_arial, font_path_dejavu]
                
                for font_path_attempt in font_load_attempts:
                    try:
                        font = ImageFont.truetype(font_path_attempt, font_size)
                        break # Font loaded successfully
                    except IOError:
                        font = None # Continue to next attempt
                
                if not font:
                    font = ImageFont.load_default() # Fallback to Pillow's default bitmap font
                    # Default font is small, consider adjusting text_position or font_size if this is often hit
                    # For now, we'll use it as is.
            except Exception as e:
                # print(f"General font loading error: {e}")
                font = ImageFont.load_default() # Ultimate fallback

            for i, img_rgba in enumerate(images_to_splice):
                img_copy = img_rgba.copy() # Work on a copy
                draw = ImageDraw.Draw(img_copy)
                text = f"图{i+1}"
                
                text_position = (10, 10)
                selected_text_color_name = self.watermark_color_var.get()
                
                if selected_text_color_name == "白色":
                    text_color = (255, 255, 255, 255) # White
                    shadow_color = (0, 0, 0, 128)     # Black shadow
                elif selected_text_color_name == "黑色":
                    text_color = (0, 0, 0, 255)       # Black
                    shadow_color = (255, 255, 255, 128) # White shadow
                else: # Default to Red
                    text_color = (255, 0, 0, 255)     # Red
                    shadow_color = (0, 0, 0, 128)     # Black shadow

                shadow_offset = 1
                # Draw shadow first
                draw.text((text_position[0]+shadow_offset, text_position[1]+shadow_offset), text, font=font, fill=shadow_color)
                # Draw main text
                draw.text(text_position, text, font=font, fill=text_color)
                
                images_with_watermark.append(img_copy)
            images_to_splice = images_with_watermark # Replace with watermarked images

        mode = self.splice_mode_var.get()
        output_image = None

        if mode == "横向拼接":
            output_image = self.splice_horizontal(images_to_splice)
        elif mode == "纵向拼接":
            output_image = self.splice_vertical(images_to_splice)
        elif mode == "2xN网格":
            if len(images_to_splice) < 1: # Allow even 1 image for 2xN, it will be a 1x1 grid
                messagebox.showwarning("提示", "2xN网格拼接至少需要1张图片。")
                return
            output_image = self.splice_grid_2xn(images_to_splice)
        else:
            messagebox.showerror("错误", "未知的拼接模式。")
            # No need to close images_to_splice here, they are copies or converted versions
            return

        if output_image:
            self.update_preview(output_image)
        
        # images_to_splice are modified copies or results of conversion, Pillow handles their memory.
        # Original images (images_to_splice_orig) are closed in the finally block.

    def splice_horizontal(self, images):
        if not images: return None
        widths, heights = zip(*(i.size for i in images))
        total_width = sum(widths)
        max_height = max(heights)
        
        # Create an RGBA canvas with a transparent background
        new_im = Image.new('RGBA', (total_width, max_height), (0, 0, 0, 0))
        
        x_offset = 0
        for im in images: # im should be RGBA here
            # Paste with alpha compositing
            new_im.paste(im, (x_offset, (max_height - im.height) // 2), im if im.mode == 'RGBA' else None)
            x_offset += im.width
        return new_im

    def splice_vertical(self, images):
        if not images: return None
        widths, heights = zip(*(i.size for i in images))
        max_width = max(widths)
        total_height = sum(heights)

        # Create an RGBA canvas with a transparent background
        new_im = Image.new('RGBA', (max_width, total_height), (0, 0, 0, 0))

        y_offset = 0
        for im in images: # im should be RGBA here
            x_pos = (max_width - im.width) // 2
            # Paste with alpha compositing
            new_im.paste(im, (x_pos, y_offset), im if im.mode == 'RGBA' else None)
            y_offset += im.height
        return new_im

    def splice_grid_2xn(self, images): # Compact "flow" layout
        if not images: return None
        num_images = len(images)
        cols = 2
        if num_images == 1:
            return images[0] # Single image, just return it

        # Calculate row structure and dimensions
        row_info = [] # Stores (actual_row_width, max_row_height, [images_in_row])
        
        current_y_offset = 0
        max_overall_width = 0
        
        i = 0
        while i < num_images:
            images_in_current_row = []
            current_row_width = 0
            current_row_max_height = 0

            # First image in the row
            img1 = images[i]
            images_in_current_row.append(img1)
            current_row_width += img1.width
            current_row_max_height = max(current_row_max_height, img1.height)
            i += 1

            # Second image in the row, if available
            if i < num_images:
                img2 = images[i]
                images_in_current_row.append(img2)
                current_row_width += img2.width # Images are side-by-side
                current_row_max_height = max(current_row_max_height, img2.height)
                i += 1
            
            row_info.append({
                "width": current_row_width,
                "height": current_row_max_height,
                "images": images_in_current_row,
                "y_start": current_y_offset
            })
            max_overall_width = max(max_overall_width, current_row_width)
            current_y_offset += current_row_max_height

        if max_overall_width == 0 or current_y_offset == 0: # Should not happen if images exist
            return None

        # Create the new image canvas with a transparent background
        total_height = current_y_offset
        new_im = Image.new('RGBA', (max_overall_width, total_height), (0, 0, 0, 0))

        # Paste images
        for row_data in row_info:
            current_x = 0
            y_pos = row_data["y_start"]
            for img_idx, im in enumerate(row_data["images"]): # im should be RGBA here
                # For this compact layout, images are simply placed one after another.
                # Vertical alignment within a row: align to the top of the row.
                # Paste with alpha compositing
                new_im.paste(im, (current_x, y_pos), im if im.mode == 'RGBA' else None)
                current_x += im.width
        
        return new_im

    def save_image(self):
        if not self.processed_image: # self.processed_image should now be RGBA
            messagebox.showwarning("提示", "没有可保存的图片。")
            return

        selected_format = self.save_format_var.get()
        
        if selected_format == "PNG":
            default_ext = ".png"
            file_types = (("PNG 文件", "*.png"), ("所有文件", "*.*"))
        elif selected_format == "JPG":
            default_ext = ".jpg"
            file_types = (("JPEG 文件", "*.jpg"), ("所有文件", "*.*"))
        else: # Should not happen with combobox
            messagebox.showerror("错误", "未知的保存格式。")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=file_types,
            title=f"保存为 {selected_format}"
        )

        if file_path:
            try:
                img_to_save = self.processed_image # This is an RGBA image

                if selected_format == "JPG":
                    # JPG doesn't support alpha, so create a white background and paste onto it.
                    if img_to_save.mode == 'RGBA':
                        # Create a new RGB image with a white background
                        background = Image.new("RGB", img_to_save.size, (255, 255, 255))
                        # Paste the RGBA image onto the white background.
                        # The alpha channel of img_to_save will be used for compositing.
                        background.paste(img_to_save, (0, 0), img_to_save)
                        img_to_save = background # Now it's an RGB image
                    elif img_to_save.mode != 'RGB': # Should not happen if processed_image is always RGBA
                        img_to_save = img_to_save.convert('RGB')
                    
                    img_to_save.save(file_path, quality=95)
                    messagebox.showinfo("成功", f"图片已保存为 JPG 到: {file_path}")

                elif selected_format == "PNG":
                    # Save directly as PNG, preserving transparency
                    img_to_save.save(file_path)
                    messagebox.showinfo("成功", f"图片已保存为 PNG 到: {file_path}")

            except Exception as e:
                messagebox.showerror("错误", f"保存图片失败: {e}")

    def copy_image(self):
        if not self.processed_image:
            messagebox.showwarning("提示", "没有可复制的图片。")
            return
        
        try:
            import io
            import win32clipboard
            
            output = io.BytesIO()
            img_to_copy = self.processed_image
            
            # For clipboard, BMP (DIB) is widely supported.
            # Ensure image is in a suitable mode for BMP (e.g., RGB, L).
            if img_to_copy.mode == 'RGBA' or img_to_copy.mode == 'P':
                 img_to_copy = img_to_copy.convert('RGB') # Convert to RGB for BMP if it has alpha or is palette-based

            img_to_copy.save(output, "BMP")
            data = output.getvalue()[14:] # Skip BMP file header for CF_DIB
            output.close()

            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            win32clipboard.CloseClipboard()
            messagebox.showinfo("成功", "图片已复制到剪贴板。")

        except ImportError:
            messagebox.showerror("错误", "复制功能需要 `pywin32` 库。\n请在命令行运行: pip install pywin32")
        except Exception as e:
            messagebox.showerror("错误", f"复制图片到剪贴板失败: {e}")


if __name__ == '__main__':
    if TkinterDnD:
        root = TkinterDnD.Tk() # Use TkinterDnD.Tk() if available
    else:
        root = tk.Tk() # Fallback to standard tk.Tk()

    app = ImageSplicerApp(root)
    
    # Ensure canvas dimensions are known before first preview update attempt
    # by forcing an update of the window geometry.
    # This is important for both preview_canvas and thumbnail_canvas
    root.update_idletasks()
    
    root.mainloop()