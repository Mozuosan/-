import pandas as pd
import codecs
import pathlib
import glob
import shutil

output_dir = "C:\\Users\\USRE NAME\\Desktop\\out1" #生成したXMLファイルの移動先フォルダ

path = pathlib.Path(".\CSV_FluxXMLConvertor")  #CSVファイル保存先フォルダのパスを指定
for pass_obj in path.iterdir():
    if pass_obj.match("*.csv"):
        df = pd.read_csv(pass_obj)
        for data in df.itertuples():
            print("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
                "<order>\n"
                "<orderInformation>\n"
                "<addresses>\n"
                "<submitter><!-- 発送者 -->\n"
                "</submitter>\n"
                "</addresses>\n"
                "</orderInformation>\n"
                "<orderItems>\n"
                "<orderItem>\n"
                "<title>" + str(data.製品番号) + "_" + str(data.製品名) + "</title> <!-- 印刷ジョブ名称 -->\n"
                "<quantity>" + str(data.数量) + "</quantity> <!-- 印刷数量 -->\n"
                "<document>\n"
                "<product>\n"
                "<name>" + str(data.仕様) + "</name> <!-- Fluxプリント商材名 -->\n"
                "<services>\n"
                "<service id=\"PAPERTYPE\">\n"
                "<option>\n"
                "<name>" + str(data.用紙銘柄) + "</name> <!-- 用紙タイプ -->\n"
                "</option>\n"
                "</service>\n"
                "</services>\n"
                "</product>\n"
                "<pageSources>\n"
                "<pageSource>\n"
                "<fileName>" + str(data.印刷ファイル名) + "</fileName> <!-- 印刷ファイル-->\n"
                "</pageSource>\n"
                "</pageSources>\n"
                "</document>\n"
                "</orderItem>\n"
                "</orderItems>\n"
                "</order>", file=codecs.open(str(data.製品番号)+'_'+str(data.製品名)+'.xml', 'w', 'utf-8'))

for file in glob.glob('./*.xml', recursive=False): #XMLファイルはpyファイルと同じフォルダに生成されるのでoutput_dirに移動する
    print(file)
    shutil.move(file, output_dir)
