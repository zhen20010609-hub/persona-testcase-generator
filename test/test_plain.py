from generator_plain import generate_plain_test_cases_excel

output_file = generate_plain_test_cases_excel(
    products="AltScoreTelco_PH",
    id_value="011115634849",
    id_type="UMID",
    cell="09206587342",
    name="aaa",
    country="PH",
    output_path="output/1-AltScoreTelco_PH-weakVerify-plain.xlsx"
)

print("明文生成成功：", output_file)