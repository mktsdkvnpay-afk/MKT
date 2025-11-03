# MKT Auto Subtitle Generator

Ứng dụng dòng lệnh giúp tự động tạo phụ đề `.srt` từ file audio hoặc video bằng mô hình [OpenAI Whisper](https://github.com/openai/whisper).

## Yêu cầu hệ thống

- Python 3.9 trở lên.
- [FFmpeg](https://ffmpeg.org/) (Whisper sử dụng FFmpeg để đọc nhiều định dạng audio/video).
- Gói `openai-whisper` (cài đặt trong bước bên dưới).

## Cài đặt

```bash
python -m venv .venv
source .venv/bin/activate  # Trên Windows: .venv\\Scripts\\activate
pip install --upgrade pip
pip install .
```

Sau khi cài đặt, lệnh `auto-sub` sẽ khả dụng trong terminal.

## Sử dụng

Tạo phụ đề tiếng Việt (Whisper tự nhận diện ngôn ngữ):

```bash
auto-sub path/to/file.mp3
```

Xuất phụ đề sang tiếng Anh:

```bash
auto-sub path/to/file.mp3 --translate
```

Chỉ định model Whisper khác (ví dụ `small`) và nơi lưu file `.srt`:

```bash
auto-sub input.mp4 --model small --output subtitles/output.srt
```

Chạy suy luận trên GPU Nvidia (yêu cầu đã cài driver/CUDA và PyTorch phù hợp):

```bash
auto-sub input.mp4 --device gpu
```

Các tuỳ chọn khác:

- `--language`: mã ngôn ngữ nguồn (ví dụ `vi`, `en`).
- `--temperature`: nhiệt độ sampling, mặc định `0.0`.
- `--beam-size`: beam search size.
- `--best-of`: số mẫu tạo trước khi chọn kết quả tốt nhất.

File `.srt` đầu ra sẽ nằm cùng thư mục với file audio/video nếu không chỉ định `--output`.

## Phát triển

Chạy ứng dụng trực tiếp từ mã nguồn mà không cần cài đặt toàn bộ gói:

```bash
PYTHONPATH=src python -m auto_sub.cli path/to/file.wav
```

## Giấy phép

Mã nguồn phát hành nội bộ cho phòng MKT.
