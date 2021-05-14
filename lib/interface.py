# -*- coding: utf-8 -*-
import traceback

import flask, time, json
from bin.printTestNo import *
from lib import tools
from config import setting
from jsonschema import validate, exceptions, ValidationError

# server = flask.Flask(__name__)

schema = {
    "type": "array",
    "items": [
        {
            "photo": {"type": "string"},
            "testno": {"type": "string"},
            "cname": {"type": "string"},
            "ename": {"type": "string"},
            "idno": {"type": "string"},
            "school": {"type": "string"},
            "examno": {"type": "string"},
            "seatno": {"type": "string"},
            "scale": {"type": "string"},
            "examDate": {"type": "string"},
            "examTime": {"type": "string"},
        }
    ]
}


# validate({"name" : "Eggs", "price" : 34.99}, schema)

def TestNoServer(server):
    @server.route('/printTestNo', methods=['post'])
    def printTestNo():
        item = flask.request.get_json()
        print("item:", item)
        # print("flag:", validate({"name": "Eggs", "price": 34.99}, schema))
        # print("flag:", validate(item, schema))
        try:
            validate(item, schema)
            # validate({"name": "Eggs", "price": 34.99}, schema)
        except ValidationError:
            res = {'msg': '注意：提交的参数不全'}  # 给用户返回的信息
            json_res = json.dumps(res, ensure_ascii=False)  # 返回结果为json格式
            res = flask.make_response(json_res)  # cookie 构造成返回结果的对象
            return res
        try:
            for i in range(0, len(item)):
                openWord(item[i]['testno'], item[i]['cname'], item[i]['ename'],
                         item[i]['idno'], item[i]['school'], item[i]['examno'],
                         item[i]['seatno'], item[i]['scale'], item[i]['examDate']+' '+item[i]['examTime'],
                         item[i]['photo'], item[i]['businessId'], item[i]['businessType'], item[i]['createdBy'])
        except Exception:
            # print("Exception:", traceback.print_exc())
            traceback.print_exc()
            res = {'msg': '注意：系统出错，请重新提交数据'}  # 给用户返回的信息
            json_res = json.dumps(res, ensure_ascii=False)  # 返回结果为json格式
            res = flask.make_response(json_res, 500)  # cookie 构造成返回结果的对象
            print("res:", res)
            return res

        res = {'msg': '生成准考证成功'}  # 给用户返回的信息
        json_res = json.dumps(res, ensure_ascii=False)  # 返回结果为json格式
        res = flask.make_response(json_res)  # cookie 构造成返回结果的对象
        # print("res:", res)
        # res.set_cookie(key, session_id, 6000)  # 最后的数字是cookie的失效时间
        return res
        # username = flask.request.values.get('username')
        # pwd = flask.request.values.get('pwd')
        # if username == 'wind' and pwd == '123456':
        #     # session_id = lib.tools.my_md5(username+time.strftime('%Y%m%d%H%M%S'))
        #     # key = 'wind_session:%s' % username
        #     # lib.tools.op_redis(key, session_id, 6000)
        #     test()
        #     res = {'session_id': username, 'error_code': 0, 'msg': '登录成功', 'login_time': time.strftime('%Y%m%d%H%M%S')}  # 给用户返回的信息
        #     json_res = json.dumps(res, ensure_ascii=False)  # 返回结果为json格式
        #     res = flask.make_response(json_res)   # cookie 构造成返回结果的对象
        #     print("res:", res)
        #     # res.set_cookie(key, session_id, 6000)  # 最后的数字是cookie的失效时间
        #     return res

        # return
