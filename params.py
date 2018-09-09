PARAMS = \
    {
        "read_params": {
            "source_name": "source.xlsx",
            "template_name": "template",     # "template1.xlsm","template2.xlsm","template3.xlsm"
            "picture_name": "curve.png",          # 因为无法自动生成形状，需要读取png
        },
        "output_params":{
            "number": 0,                # 0 表示excel里面所有行都要读出来，其他数字表示强行限制读出的数量
            "output_name": "result",    # 表示生成文档的名字，如果是"output"，就是"output1.docx","output2.docx" ......
        },
        "output_dir": "结果",           # 输出文件夹文件名
    }