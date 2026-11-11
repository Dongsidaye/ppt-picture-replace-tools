# Offline OCR assets

The optional OCR feature uses RapidOcrNet 4.2.0 and PaddleOCR PP-OCRv5 models.
Model files and the Chinese recognition dictionary are downloaded only while
building the add-in. Installed copies run locally without network access.

The model URLs and SHA-256 values are pinned in models.json. Detection,
classification and recognition model hashes were verified against RapidAI's
official default_models.yaml. The dictionary comes from PaddleOCR v3.2.0 and
is paired with the Chinese PP-OCRv5 mobile recognizer.

PaddleOCR and RapidOCR models are distributed under Apache License 2.0.
The original license text is supplied as LICENSE-PaddleOCR.txt.

- https://github.com/RapidAI/RapidOCR
- https://github.com/RapidAI/RapidOCR/blob/main/python/rapidocr/default_models.yaml
- https://github.com/PaddlePaddle/PaddleOCR/tree/v3.2.0
- https://github.com/BobLd/RapidOcrNet

The OCR implementation uses RapidOcrNet 4.2.0 (Apache-2.0), ONNX Runtime 1.29.0
(MIT), SkiaSharp 3.119.4 (MIT), Svg.Skia 4.9.1 (MIT), HarfBuzzSharp 8.3.1.3 (MIT),
Clipper2 2.0.0 (Boost Software License 1.0), and ExCSS 4.3.1 (MIT).
Svg.Skia's Svg.Custom dependency is distributed under Microsoft Public License
(MS-PL). Its original license is included separately. Svg.Animation, Svg.Model,
Svg.SceneGraph and ShimSkiaSharp 4.9.1 use the Svg.Skia repository's MIT license.
The code-licenses folder supplies the corresponding original texts and notices.

OCR can split a long label into multiple regions, misread characters, or find
text-like drawing geometry. Each region includes a confidence estimate and
four corners for review before editable text is inserted.
