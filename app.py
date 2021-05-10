# -*- coding: utf-8 -*-
import config.setting as config
import sys

# import redis
from lib.interface import *
from flask_cors import CORS

print(sys.path)


server = flask.Flask(__name__)
CORS(server, resources=r'/*')

server.config.from_object(config)

# @server.route('/index', methods=['get'])
# def index():
#     res = {'msg': '这是一个接口', 'msg_code': 0}
#     return json.dumps(res,ensure_ascii=False)

# # 操作redis
# def op_redis(k, v=None):
#     print("REDIS_INFO:", server.config['REDIS_INFO'])
#     r = redis.Redis(server.config['REDIS_INFO'])
#     return r


TestNoServer(server)


# if __name__ == '__main__':
# server.run(port=8999,debug=True)
# server.run(port=8999,host='172.28.28.86')
server.run(port=8999,host='172.17.0.2')
