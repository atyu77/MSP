import datetime
import logging
import os.path
import pyodbc
import requests
import time
import wget
import xml.etree.cElementTree as ET
import zipfile_deflate64 as zipfile


from bs4 import BeautifulSoup
from logging.handlers import RotatingFileHandler

logger = logging.getLogger(__name__)

BASE_DIR = os.path.abspath(os.curdir)
url = "https://www.nalog.gov.ru/opendata/7707329152-rsmp/"
Dir = "C:\\Scripts\\MSP\\Files"
conn = pyodbc.connect("Driver={SQL Server};"
                      "Server=VABS2;"
                      "Database=MSP;"
                      "UID=loader;"
                      "PWD=123456"
                      )

conn1 = pyodbc.connect("Driver={SQL Server};"
                       "Server=ABS3;"
                       "Database=ALORBANK;"
                       "UID=loader;"
                       "PWD=123456"
                       )
RETRY_PERIOD = 7200


def get_text(node) -> str:
    return node.get_text(strip=True) if node else ''


def parseHTML(url):
    response = requests.get(url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "lxml")
        table = soup.find("table", class_="border_table")
        for row in table.find_all("tr"):
            tds = row.select("td")
            if len(tds) > 0:
                td_number, td_name, td_value = tds
                if get_text(td_number) == "12":
                    date_last_change = get_text(td_value)
                if get_text(td_number) == "8":
                    link = get_text(td_value)
    return {"date_last_change": date_last_change,
            "link": link}


def download(link, Dir, dtMSP):
    logger.info(f'Скачиваем реестр МСП за дату {dtMSP} в {Dir}')
    wget.download(link, Dir+"\\data_files.zip")


def unzip(Dir):
    logger.info(f'Распаковываем архив в {Dir}')
    print(f'Распаковываем архив в {Dir}')
    try:
        MSPzip = zipfile.ZipFile(Dir+"\\data_files.zip")
        MSPzip.extractall(Dir)
    except Exception as error:
        logger.error(str(error))
        print(str(error))
    #    subprocess.Popen(["7z", "e", Dir+"\\data_files.zip", f"-o{Dir}", "-y"])


def deleteFiles(Dir):
    print("Очищаем папку")
    logger.info("Очищаем папку")
    for f in os.listdir(Dir):
        os.remove(os.path.join(Dir, f))


def ExecuteSQL(cursor, SQL):
    res = 0
    try:
        retval = cursor.execute(SQL)
        res = len(retval.fetchall())
        cursor.connection.commit()
    except Exception as error:
        logger.error(str(error))
        print(str(error))
        res = -1
    return res


def ResultSQL(cursor, SQL):
    res = 0
    try:
        rows = cursor.execute(SQL).fetchall()
        res = rows[0][0]
        cursor.commit()
    except Exception as error:
        print(str(error))
        logger.error(str(error))
        res = -1
    return res


def parseXML(cursor, Dir, file):
    count = 0
    listMSP = []
    listMSPFile = []
    tree = ET.parse(Dir+"\\"+file)
    docs = tree.findall("Документ")
    try:
        for doc in docs:
            if doc.find("ИПВклМСП") is not None:
                Inn_IP = doc.find("ИПВклМСП").attrib["ИННФЛ"]
            else:
                Inn_IP = ""
            listMSP.append(Inn_IP)
            if doc.find("ОргВклМСП") is not None:
                Inn_UL = doc.find("ОргВклМСП").attrib["ИННЮЛ"]
            else:
                Inn_UL = ""
            listMSP.append(Inn_UL)
            listMSP.append(doc.attrib["ДатаВклМСП"])
            listMSP.append(doc.attrib["КатСубМСП"])
            listMSP.append(file)
            count = count+1
            listMSPFile.append(listMSP)
            listMSP = []
        cursor.executemany("Insert into tMSP VALUES(?,?,?,?,?)", listMSPFile)
        cursor.commit()
    except Exception as error:
        print(str(error))
        logger.error(str(error))
        count = -1
    return count


def MSP_parse():
    count = 0
    countIt = 0
    countF = 0
    dtMSP = f'10.{str(format(datetime.datetime.now().month, "02"))}.{str(datetime.datetime.now().year)}'
    try:
        cursor = conn.cursor()
        cursor.fast_executemany = True
        SQL = "select 1 from tDt_MSP where Date  ='" + dtMSP+"'"
        res = ExecuteSQL(cursor, SQL)
        if res == 0:
            print(f'Реестр МСП за дату {dtMSP} еще не скачивался')
            logger.info(f'Реестр МСП за дату {dtMSP} еще не скачивался')
            dict = parseHTML(url)
            if dict["date_last_change"] == dtMSP:
                try:
                    deleteFiles(Dir)
                    download(dict["link"], Dir, dtMSP)
                    if os.path.exists(Dir+"\\data_files.zip") is True:
                        try:
                            unzip(Dir)
                            os.remove(Dir+"\\data_files.zip")
                        except Exception as error:
                            print(str(error))
                            logger.error(str(error))
                except Exception as error:
                    print(str(error))
                    logger.error(str(error))
                print("Чистим таблицу tMSP")
                logger.info("Чистим таблицу tMSP")
                SQL = """
                         SET NOCOUNT ON
                         SET ANSI_WARNINGS OFF
                         declare @c int
                         select @c = count(1) from tMSP
                         if @c>0 delete from tMSP
                         select @c as c"""
                countdel = ResultSQL(conn, SQL)
                print(f'Удалено {countdel} записей из tMSP')
                logger.info(f'Удалено {countdel} записей из tMSP')
                print("Начинаем обработку xml файлов")
                logger.info("Начинаем обработку xml файлов")
                for root, dirs,  files in os.walk(Dir):
                    for file in files:
                        if (file.endswith(".xml")):
                            countF = countF + 1
                            count = parseXML(cursor, Dir, file)
                        if count > 0:
                            countIt = countIt+count
                        else:
                            break
                if countIt > 0:
                    print(f'Добавлено {countIt} записей в tMSP из {countF} файлов')
                    logger.info(f'Добавлено {countIt} записей в tMSP из {countF} файлов')
                    cursorABS3 = conn1.cursor()
                    SQL = f"SET NOCOUNT ON declare @dt smalldatetime select @dt = convert(date, '{dtMSP}', 104) exec MSPInst @date = @dt"
                    try:
                        cursorABS3.execute(SQL)
                        cursorABS3.commit()
                        SQL = f"insert into tDt_MSP VALUES ('{dtMSP}') select count(1) from tDt_MSP where date = '{dtMSP}'"
                        countDt = ResultSQL(conn, SQL)
                        print(f"Добавлена {countDt} записей в tDt_MSP")
                        logger.info(f"Добавлена {countDt} записей в tDt_MSP")
                    except Exception as error:
                        print(str(error))
                        logger.error(str(error))
            else:
                print(f'Реестр за {dtMSP} не найден на сайте налоговой. Последнее обновление {dict["date_last_change"]}.')
                logger.info(f'Реестр за {dtMSP} не найден на сайте налоговой. Последнее обновление {dict["date_last_change"]}.')
        if res > 0:
            print(f'Реестр МСП за дату {dtMSP} ранее был скачан')
            logger.info(f'Реестр МСП за дату {dtMSP} ранее был скачан')
        if res < 0:
            logger.info("Подключение к БД прошло неуспешно")
    except Exception as error:
        print(str(error))
        logger.error(str(error))


def main():
    while True:
        MSP_parse()
        time.sleep(RETRY_PERIOD)


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO,
                        format=('%(asctime)s [%(levelname)s] - '
                                '(%(filename)s).%(funcName)s:'
                                '%(lineno)d - %(message)s'
                                ),
                        handlers=[
                            RotatingFileHandler(
                                f'{BASE_DIR}/LOG/output.log',
                                mode="a",
                                maxBytes=4000000,
                                backupCount=1000,
                                encoding=None,
                                delay=False)
                                ]
                        )
    logger = logging.getLogger('Logs')
    logger.setLevel(logging.DEBUG)
    main()
