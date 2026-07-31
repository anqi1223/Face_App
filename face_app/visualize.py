"""可视化 - 在人脸图片上绘制检测框、关键点和标注信息"""
import cv2


def draw_face_annotations(img, faces, names=None, scores=None):
    """
    在图片上绘制人脸检测框、5 个关键点和标注文字。
    返回标注后的图片副本（不修改原图）。
    """
    annotated = img.copy()

    for i, face in enumerate(faces):
        # --- 1. 边界框 ---
        bbox = face.bbox.astype(int)
        x1, y1, x2, y2 = bbox

        if names and i < len(names):
            color = (0, 255, 0) if names[i] != "Unknown" else (0, 0, 255)   # 绿 / 红
        else:
            color = (255, 0, 0)                                              # 蓝（录入模式）

        cv2.rectangle(annotated, (x1, y1), (x2, y2), color, 2)

        # --- 2. 5 个关键点（双眼、鼻尖、左右嘴角）---
        if hasattr(face, 'kps') and face.kps is not None:
            for kp in face.kps:
                cv2.circle(annotated, (int(kp[0]), int(kp[1])), 2, (0, 255, 255), -1)

        # --- 3. 标注文字 ---
        if names and i < len(names):
            label = f"{names[i]} ({scores[i]:.2f})" if scores and i < len(scores) else names[i]
        else:
            label = f"Face {i + 1}"

        font = cv2.FONT_HERSHEY_SIMPLEX
        (tw, th), _ = cv2.getTextSize(label, font, 0.5, 1)
        cv2.rectangle(annotated, (x1, y1 - th - 8), (x1 + tw, y1), color, -1)
        cv2.putText(annotated, label, (x1, y1 - 4), font, 0.5, (255, 255, 255), 1)

    return annotated
