import pymysql
import uuid
import time

class Repository:
    def __init__(self):
        try:
            db = pymysql.Connect(host="rm-wz92099r66qstfwr0lo.mysql.rds.aliyuncs.com", port=3306, user="huawenshuiping", password="huawenshuiping123456", database="hwcsWeb", charset='utf8')
            # db = pymysql.Connect(host="172.28.28.86", port=3306, user="root", password="san1546", database="hwcsWeb", charset='utf8')
            self.db = db
        except Exception as e:
            print(e)

    def saveExamineeCardAtt(self, name, path, size, ext, business_id, business_type, created_by):
        # 使用cursor()方法获取操作游标
        cursor = self.db.cursor()
        attachmentid = str(uuid.uuid4()).replace("-", "")
        nowtime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        # SQL 插入语句
        sql = "INSERT INTO attachments(id,name,path,ext,size,business_id,business_type,created_at,updated_at,created_by)\
                VALUES ('%s', '%s',  '%s',  '%s',  '%s',  '%s',  '%s',  '%s',  '%s',  '%s')" % \
            (attachmentid, name, path, ext, size, business_id, business_type, nowtime, nowtime, created_by)
        try:
            # 执行sql语句
            cursor.execute(sql)
            # 提交到数据库执行
            self.db.commit()
        except Exception as e:
            print(e)
            # 如果发生错误则回滚
            self.db.rollback()

