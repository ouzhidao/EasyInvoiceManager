# PaddleOCR修复说明

## 问题现象

如果软件运行时出现以下错误：
```
(Unimplemented) ConvertPirAttribute2RuntimeAttribute not support [pir::ArrayAttribute<pir::DoubleAttribute>]
```

这是PaddlePaddle 3.3.0版本与部分CPU/系统的兼容性问题。

## 解决方案（二选一）

### 方案1：配置百度API（推荐，最简单）

1. 打开软件，点击"配置API"按钮
2. 输入百度OCR的密钥（App ID、API Key、Secret Key）
3. 配置后软件会优先使用API，不再依赖本地PaddleOCR

百度OCR申请地址：https://ai.baidu.com/tech/ocr
每天有500次免费调用额度。

### 方案2：降级PaddlePaddle版本

在命令行执行以下命令：

```bash
pip uninstall paddlepaddle -y
pip install paddlepaddle==2.6.1 -i https://pypi.tuna.tsinghua.edu.cn/simple
```

降级后重启软件即可。

## 验证修复

修复后，可以在命令行测试PaddleOCR是否正常：

```bash
python -c "from paddleocr import PaddleOCR; ocr = PaddleOCR(lang='ch'); print('PaddleOCR初始化成功')"
```

如果没有报错，说明修复成功。

## 其他说明

- 软件逻辑：配置了API时优先使用API，未配置时使用本地PaddleOCR
- 如果API额度用完，会自动降级到本地PaddleOCR
- 建议同时配置API和保持本地OCR可用，以确保最佳体验
