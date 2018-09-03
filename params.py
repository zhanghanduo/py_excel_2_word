PARAMS = \
    {
        "read_params": {
            "excel_name": "test.xlsx",
            "template_name": "template_0.docx",
        },
        "output_params":{
            "number": 0,                # 0 表示excel里面所有行都要读出来，其他数字表示强行限制读出的数量
            "output_name": "result",    # 表示生成文档的名字，如果是"output"，就是"output1.docx","output2.docx" ......
        },
        "date_params":  0,              # 0 表示不添加日期在文档底部，自己会手动添加，1 表示添加当前日期在文档底部
        "random_params":0,
        "output_dir": "output",         # 输出文件夹文件名
    }