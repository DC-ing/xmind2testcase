# 使用官方 Python 镜像作为基础镜像
FROM alibaba-cloud-linux-3-registry.cn-hangzhou.cr.aliyuncs.com/alinux3/python:3.11.1

# 设置 Flask 应用的环境变量
ENV FLASK_APP=webtool/application.py
ENV FLASK_RUN_HOST=0.0.0.0
ENV FLASK_PORT=5001
ENV DOCKER_WORK_DIR=/xmind2testcase

# 设置工作目录，切换到容器目录
WORKDIR ${DOCKER_WORK_DIR}

# 复制当前目录下的所有文件到容器的工作目录
COPY . .

# 安装项目依赖
RUN pip install --no-cache-dir -r requirements.txt


# 暴露 FLask 端口
EXPOSE ${FLASK_PORT}

# 运行 Flask 应用
CMD [ "python", "webtool/application.py" ]
