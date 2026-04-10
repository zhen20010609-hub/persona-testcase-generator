from generator_MD5 import generate_md5_test_cases_excel

output_file = generate_md5_test_cases_excel(
    products="AltScoreTelco_PH",
    id_value="011115634849",
    id_type="UMID",
    cell="09206587342",
    name="aaa",
    country="PH",
    output_path="output/1-AltScoreTelco_PH-weakVerify-md5.xlsx"
)

print("MD5生成成功：", output_file)