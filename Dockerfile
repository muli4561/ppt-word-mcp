# ppt-word-gen 镜像：默认联网安装；企业构建脚本可切换为本地 wheel 离线安装。
FROM python:3.12-slim@sha256:2c941e860699f878900b0edc2403613c234d4b32eda3cc9fa7036991a2a63c4a

WORKDIR /app

# 根目录 requirements.txt 是本机与容器共用的唯一清单；.docker-wheels 只保存安装包。
COPY requirements.txt /tmp/requirements.txt
COPY .docker-wheels /wheels
COPY .word-wheels /word-wheels
COPY .mcp-wheels /mcp-wheels
ARG OFFLINE_INSTALL=0
RUN if [ "$OFFLINE_INSTALL" = "1" ]; then \
        pip install --no-index --find-links /wheels --find-links /word-wheels --find-links /mcp-wheels -r /tmp/requirements.txt; \
    else \
        pip install --no-cache-dir --find-links /wheels --find-links /word-wheels --find-links /mcp-wheels -r /tmp/requirements.txt; \
    fi \
    && rm -rf /wheels /word-wheels /mcp-wheels /tmp/requirements.txt

# Python 包：单一生成服务
COPY ppt_word_gen /app/ppt_word_gen
COPY skills /app/skills
COPY static /app/static
COPY assets /app/assets

# ppt-master skill 仓库：默认挂载到 /ppt-master（compose 里通过 volume 挂载）
ENV PPTMASTER_ROOT=/ppt-master

EXPOSE 8000
CMD ["uvicorn", "ppt_word_gen.app:app", "--host", "0.0.0.0", "--port", "8000"]
