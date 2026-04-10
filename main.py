import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from generator_plain import generate_plain_test_cases_excel, generate_test_cases as generate_plain_test_cases
from generator_MD5 import generate_md5_test_cases_excel, generate_test_cases as generate_md5_test_cases


RESERVED_KEYS = {"products", "id", "id_type", "cell", "name", "country"}


def choose_output_dir():
    folder_path = filedialog.askdirectory(title="选择输出文件夹")
    if folder_path:
        entry_output_dir.delete(0, tk.END)
        entry_output_dir.insert(0, folder_path)


def split_and_clean_csv(text):
    """
    按英文逗号分隔，并自动去掉每项首尾空格，忽略空项
    例如: " gaid, email , imei , " -> ["gaid", "email", "imei"]
    """
    return [item.strip() for item in text.split(",") if item.strip()]


def toggle_extra_fields():
    refresh_layout()


def toggle_id_type(*args):
    """
    只有 country 严格等于 PH 时显示 id_type
    """
    if entry_country.get().strip() != "PH":
        entry_id_type.delete(0, tk.END)
    refresh_layout()


def build_extra_fields():
    if need_extra_key.get() != "yes":
        return {}

    raw_keys = entry_extra_key.get().strip()
    raw_values = entry_extra_value.get().strip()

    if not raw_keys:
        raise ValueError("请选择了需要扩展 key，但扩展 key 为空")
    if not raw_values:
        raise ValueError("请选择了需要扩展 key，但扩展 value 为空")

    key_list = split_and_clean_csv(raw_keys)
    value_list = split_and_clean_csv(raw_values)

    if not key_list:
        raise ValueError("扩展 key 解析后为空，请检查输入")
    if not value_list:
        raise ValueError("扩展 value 解析后为空，请检查输入")

    if len(key_list) != len(value_list):
        raise ValueError(
            f"扩展 key 和扩展 value 数量不一致：\n"
            f"key 数量 = {len(key_list)}\n"
            f"value 数量 = {len(value_list)}\n\n"
            f"请使用英文逗号分隔，并保证一一对应"
        )

    invalid_key_chars = {'"', "'", "\t", "\n", "\r"}

    extra_fields = {}
    seen_keys = set()

    for key, value in zip(key_list, value_list):
        if key in RESERVED_KEYS:
            raise ValueError(f"扩展 key 不能与固定字段重名：{key}")

        if any(ch in key for ch in invalid_key_chars):
            raise ValueError(f'扩展 key "{key}" 不能包含引号或换行')

        if " " in key:
            raise ValueError(f'扩展 key "{key}" 不能包含空格')

        if key in seen_keys:
            raise ValueError(f'扩展 key 存在重复项：{key}')

        seen_keys.add(key)
        extra_fields[key] = value

    return extra_fields


def collect_form_data(validate_output_filename=True):
    """
    统一收集并校验表单数据
    validate_output_filename:
        - True: 用于正式生成，需要校验输出文件名
        - False: 用于预览，可不校验输出文件名
    """
    products = entry_products.get().strip()
    country = entry_country.get().strip()
    id_value = entry_id.get().strip()
    id_type = entry_id_type.get().strip() if country == "PH" else ""
    cell = entry_cell.get().strip()
    name = entry_name.get().strip()
    output_filename = entry_output_filename.get().strip()
    output_dir = entry_output_dir.get().strip()
    encrypt_mode = encrypt_type.get()

    if not products:
        raise ValueError("products 不能为空")

    if not country:
        raise ValueError("country 不能为空")

    if country == "PH" and not id_type:
        raise ValueError("country=PH 时，id_type 不能为空")

    if not id_value:
        raise ValueError("id 不能为空")

    if not cell:
        raise ValueError("cell 不能为空")

    if not name:
        raise ValueError("name 不能为空")

    if validate_output_filename:
        if not output_filename:
            raise ValueError("输出文件名不能为空")

        invalid_chars = r'\/:*?"<>|'
        if any(c in output_filename for c in invalid_chars):
            raise ValueError('输出文件名不能包含以下字符：\\ / : * ? " < > |')

    extra_fields = build_extra_fields()

    return {
        "products": products,
        "country": country,
        "id_value": id_value,
        "id_type": id_type,
        "cell": cell,
        "name": name,
        "output_filename": output_filename,
        "output_dir": output_dir,
        "encrypt_mode": encrypt_mode,
        "extra_fields": extra_fields
    }


def get_preview_data():
    data = collect_form_data(validate_output_filename=False)

    if data["encrypt_mode"] == "plain":
        json_data, descriptions = generate_plain_test_cases(
            products=data["products"],
            id_value=data["id_value"],
            id_type=data["id_type"],
            cell=data["cell"],
            name=data["name"],
            country=data["country"],
            extra_fields=data["extra_fields"]
        )
    elif data["encrypt_mode"] == "md5":
        json_data, descriptions = generate_md5_test_cases(
            products=data["products"],
            id_value=data["id_value"],
            id_type=data["id_type"],
            cell=data["cell"],
            name=data["name"],
            country=data["country"],
            extra_fields=data["extra_fields"]
        )
    else:
        raise ValueError("请选择正确的加密方式")

    return json_data, descriptions


def show_preview_window():
    try:
        json_data, descriptions = get_preview_data()
    except ValueError as e:
        messagebox.showerror("错误", str(e))
        return
    except Exception as e:
        messagebox.showerror("错误", f"预览失败：\n{e}")
        return

    preview_win = tk.Toplevel(root)
    preview_win.title("测试用例预览")
    preview_win.geometry("1100x720")
    preview_win.minsize(900, 600)

    top_frame = tk.Frame(preview_win)
    top_frame.pack(fill="x", padx=10, pady=8)

    summary_label = tk.Label(
        top_frame,
        text=f"本次共生成 {len(json_data)} 条测试用例",
        font=("微软雅黑", 10, "bold")
    )
    summary_label.pack(side="left")

    # 上半部分：全部用例
    main_frame = tk.Frame(preview_win)
    main_frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))

    tk.Label(main_frame, text="全部用例：", anchor="w").pack(anchor="w")

    tree_container = tk.Frame(main_frame)
    tree_container.pack(fill="both", expand=True)

    columns = ("num", "desc", "req_data")
    tree = ttk.Treeview(tree_container, columns=columns, show="headings")

    tree.heading("num", text="num")
    tree.heading("desc", text="Description")
    tree.heading("req_data", text="req_data")

    # 宽度调回更接近第一版
    tree.column("num", width=70, anchor="center", stretch=False)
    tree.column("desc", width=320, anchor="w", stretch=False)
    tree.column("req_data", width=680, anchor="w", stretch=False)

    tree_vsb = ttk.Scrollbar(tree_container, orient="vertical", command=tree.yview)
    tree_hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=tree_vsb.set, xscrollcommand=tree_hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    tree_vsb.grid(row=0, column=1, sticky="ns")
    tree_hsb.grid(row=1, column=0, sticky="ew")

    tree_container.rowconfigure(0, weight=1)
    tree_container.columnconfigure(0, weight=1)

    for i, (desc, req) in enumerate(zip(descriptions, json_data), start=1):
        tree.insert("", "end", values=(i, desc, req))

    # 下半部分：选中行详情
    detail_frame = tk.Frame(preview_win)
    detail_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    tk.Label(detail_frame, text="选中行详情：", anchor="w").pack(anchor="w")

    detail_container = tk.Frame(detail_frame)
    detail_container.pack(fill="both", expand=True)

    detail_text = tk.Text(
        detail_container,
        height=10,
        wrap="none",   # 不自动换行，支持横向滚动
        undo=False
    )

    detail_vsb = ttk.Scrollbar(detail_container, orient="vertical", command=detail_text.yview)
    detail_hsb = ttk.Scrollbar(detail_container, orient="horizontal", command=detail_text.xview)

    detail_text.configure(
        yscrollcommand=detail_vsb.set,
        xscrollcommand=detail_hsb.set
    )

    detail_text.grid(row=0, column=0, sticky="nsew")
    detail_vsb.grid(row=0, column=1, sticky="ns")
    detail_hsb.grid(row=1, column=0, sticky="ew")

    detail_container.rowconfigure(0, weight=1)
    detail_container.columnconfigure(0, weight=1)

    def on_tree_select(event):
        selected = tree.selection()
        if not selected:
            return

        item = tree.item(selected[0], "values")
        if not item:
            return

        num, desc, req_data = item
        detail_text.delete("1.0", tk.END)

        try:
            pretty_json = json.dumps(json.loads(req_data), ensure_ascii=False, indent=2)
        except Exception:
            pretty_json = req_data

        content = f"序号：{num}\n\nDescription：\n{desc}\n\nreq_data：\n{pretty_json}"
        detail_text.insert(tk.END, content)

        detail_text.xview_moveto(0)
        detail_text.yview_moveto(0)

    tree.bind("<<TreeviewSelect>>", on_tree_select)

    if tree.get_children():
        first_item = tree.get_children()[0]
        tree.selection_set(first_item)
        tree.focus(first_item)
        tree.event_generate("<<TreeviewSelect>>")


def on_generate():
    try:
        data = collect_form_data(validate_output_filename=True)
    except ValueError as e:
        messagebox.showerror("错误", str(e))
        return

    output_dir = data["output_dir"]
    if not output_dir:
        output_dir = os.path.join(os.getcwd(), "output")
        os.makedirs(output_dir, exist_ok=True)
        entry_output_dir.delete(0, tk.END)
        entry_output_dir.insert(0, output_dir)

    output_path = os.path.join(output_dir, f"{data['output_filename']}.xlsx")

    if os.path.exists(output_path):
        confirm = messagebox.askyesno(
            "确认覆盖",
            f"文件已存在：\n{output_path}\n\n文件名相同会覆盖上一次生成的文件，是否继续？"
        )
        if not confirm:
            return

    try:
        if data["encrypt_mode"] == "plain":
            result_file = generate_plain_test_cases_excel(
                products=data["products"],
                id_value=data["id_value"],
                id_type=data["id_type"],
                cell=data["cell"],
                name=data["name"],
                country=data["country"],
                output_path=output_path,
                extra_fields=data["extra_fields"]
            )
        elif data["encrypt_mode"] == "md5":
            result_file = generate_md5_test_cases_excel(
                products=data["products"],
                id_value=data["id_value"],
                id_type=data["id_type"],
                cell=data["cell"],
                name=data["name"],
                country=data["country"],
                output_path=output_path,
                extra_fields=data["extra_fields"]
            )
        else:
            messagebox.showerror("错误", "请选择正确的加密方式")
            return

        messagebox.showinfo("成功", f"测试用例已生成：\n{result_file}")

    except Exception as e:
        messagebox.showerror("错误", f"生成失败：\n{e}")


def refresh_layout():
    """
    动态重排行布局：
    - country != PH 时不显示 id_type，下面控件自动上移
    - 需要扩展 key 时才显示扩展区域
    """
    for widget in form_frame.grid_slaves():
        widget.grid_forget()

    current_row = 0

    label_products.grid(row=current_row, column=0, pady=8, padx=5)
    entry_products.grid(row=current_row, column=1, pady=8, padx=5)
    current_row += 1

    label_country.grid(row=current_row, column=0, pady=8, padx=5)
    entry_country.grid(row=current_row, column=1, pady=8, padx=5)
    current_row += 1

    if entry_country.get().strip() == "PH":
        label_id_type.grid(row=current_row, column=0, pady=8, padx=5)
        entry_id_type.grid(row=current_row, column=1, pady=8, padx=5)
        current_row += 1

    label_id.grid(row=current_row, column=0, pady=8, padx=5)
    entry_id.grid(row=current_row, column=1, pady=8, padx=5)
    current_row += 1

    label_cell.grid(row=current_row, column=0, pady=8, padx=5)
    entry_cell.grid(row=current_row, column=1, pady=8, padx=5)
    current_row += 1

    label_name.grid(row=current_row, column=0, pady=8, padx=5)
    entry_name.grid(row=current_row, column=1, pady=8, padx=5)
    current_row += 1

    label_need_extra.grid(row=current_row, column=0, pady=8, padx=5)
    need_extra_frame.grid(row=current_row, column=1, pady=8, padx=5, sticky="w")
    current_row += 1

    if need_extra_key.get() == "yes":
        label_extra_key.grid(row=current_row, column=0, pady=8, padx=5)
        entry_extra_key.grid(row=current_row, column=1, pady=8, padx=5)
        current_row += 1

        label_extra_value.grid(row=current_row, column=0, pady=8, padx=5)
        entry_extra_value.grid(row=current_row, column=1, pady=8, padx=5)
        current_row += 1
    else:
        entry_extra_key.delete(0, tk.END)
        entry_extra_value.delete(0, tk.END)

    label_encrypt.grid(row=current_row, column=0, pady=8, padx=5)
    encrypt_frame.grid(row=current_row, column=1, pady=8, padx=5, sticky="w")
    current_row += 1

    label_output_filename.grid(row=current_row, column=0, pady=8, padx=5)
    entry_output_filename.grid(row=current_row, column=1, pady=8, padx=5)
    current_row += 1

    label_output_dir.grid(row=current_row, column=0, pady=8, padx=5)
    entry_output_dir.grid(row=current_row, column=1, pady=8, padx=5)
    btn_choose.grid(row=current_row, column=2, pady=8, padx=5)


root = tk.Tk()
root.title("persona基础格式校验用例生成工具")
root.geometry("760x700")
root.resizable(False, False)

title_label = tk.Label(root, text="基础格式校验用例生成工具", font=("微软雅黑", 14, "bold"))
title_label.pack(pady=15)

form_frame = tk.Frame(root)
form_frame.pack(padx=20, pady=10, fill="x")

label_products = tk.Label(form_frame, text="products", width=18, anchor="e")
entry_products = tk.Entry(form_frame, width=52)
entry_products.insert(0, "")

label_country = tk.Label(form_frame, text="country", width=18, anchor="e")
entry_country = tk.Entry(form_frame, width=52)
entry_country.insert(0, "")
entry_country.bind("<KeyRelease>", toggle_id_type)
entry_country.bind("<FocusOut>", toggle_id_type)

label_id_type = tk.Label(form_frame, text="id_type", width=18, anchor="e")
entry_id_type = tk.Entry(form_frame, width=52)
entry_id_type.insert(0, "")

label_id = tk.Label(form_frame, text="id", width=18, anchor="e")
entry_id = tk.Entry(form_frame, width=52)
entry_id.insert(0, "")

label_cell = tk.Label(form_frame, text="cell", width=18, anchor="e")
entry_cell = tk.Entry(form_frame, width=52)
entry_cell.insert(0, "")

label_name = tk.Label(form_frame, text="name", width=18, anchor="e")
entry_name = tk.Entry(form_frame, width=52)
entry_name.insert(0, "")

label_need_extra = tk.Label(form_frame, text="是否需要扩展key", width=18, anchor="e")
need_extra_key = tk.StringVar(value="no")
need_extra_frame = tk.Frame(form_frame)

tk.Radiobutton(
    need_extra_frame,
    text="否",
    variable=need_extra_key,
    value="no",
    command=toggle_extra_fields
).pack(side="left", padx=(0, 15))

tk.Radiobutton(
    need_extra_frame,
    text="是",
    variable=need_extra_key,
    value="yes",
    command=toggle_extra_fields
).pack(side="left")

label_extra_key = tk.Label(form_frame, text="扩展 key", width=18, anchor="e")
entry_extra_key = tk.Entry(form_frame, width=52)
entry_extra_key.insert(0, "")

label_extra_value = tk.Label(form_frame, text="扩展 value", width=18, anchor="e")
entry_extra_value = tk.Entry(form_frame, width=52)
entry_extra_value.insert(0, "")

label_encrypt = tk.Label(form_frame, text="加密方式", width=18, anchor="e")
encrypt_type = tk.StringVar(value="plain")
encrypt_frame = tk.Frame(form_frame)

tk.Radiobutton(encrypt_frame, text="明文", variable=encrypt_type, value="plain").pack(side="left", padx=(0, 15))
tk.Radiobutton(encrypt_frame, text="MD5", variable=encrypt_type, value="md5").pack(side="left")

label_output_filename = tk.Label(form_frame, text="输出文件名", width=18, anchor="e")
entry_output_filename = tk.Entry(form_frame, width=52)
entry_output_filename.insert(0, "")

label_output_dir = tk.Label(form_frame, text="输出文件夹", width=18, anchor="e")
entry_output_dir = tk.Entry(form_frame, width=52)
btn_choose = tk.Button(form_frame, text="选择", width=8, command=choose_output_dir)

tip_label = tk.Label(
    root,
    text=(
        "说明：\n"
        "1. 只有 country = PH 时才显示 id_type，且必须填写；\n"
        "2. 默认不启用扩展 key；\n"
        "3. 多个扩展 key / value 请用英文逗号分隔，系统会自动忽略每项前后空格，并按顺序一一对应；\n"
        "4. 可点击“预览测试用例”查看本次将生成的内容。"
    ),
    fg="gray",
    justify="left"
)
tip_label.pack(pady=5)

button_frame = tk.Frame(root)
button_frame.pack(pady=25)

btn_preview = tk.Button(
    button_frame,
    text="预览测试用例",
    width=18,
    height=2,
    command=show_preview_window
)
btn_preview.pack(side="left", padx=10)

btn_generate = tk.Button(
    button_frame,
    text="生成测试用例",
    width=18,
    height=2,
    command=on_generate
)
btn_generate.pack(side="left", padx=10)

refresh_layout()

root.mainloop()