"""Command line interface for the automatic subtitle generator."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable, Optional

from .transcriber import TranscriptionOptions, transcribe_to_srt


def _default_output_path(audio_path: Path) -> Path:
    if audio_path.suffix.lower() == ".srt":
        return audio_path
    return audio_path.with_suffix(".srt")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="auto-sub",
        description="Generate subtitle (.srt) files from audio using OpenAI Whisper.",
    )
    parser.add_argument("audio", type=Path, help="Đường dẫn file audio/video cần tạo phụ đề")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Đường dẫn file .srt xuất ra (mặc định cùng tên với file audio)",
    )
    parser.add_argument(
        "-m",
        "--model",
        default="base",
        help="Tên model Whisper (tiny, base, small, medium, large)",
    )
    parser.add_argument(
        "-l",
        "--language",
        help="Mã ngôn ngữ nguồn (vd: vi, en). Để trống để Whisper tự nhận diện",
    )
    parser.add_argument(
        "--translate",
        action="store_true",
        help="Dịch phụ đề sang tiếng Anh (sử dụng task=translate)",
    )
    parser.add_argument(
        "--temperature",
        type=float,
        default=0.0,
        help="Nhiệt độ sampling (mặc định 0)",
    )
    parser.add_argument(
        "--beam-size",
        type=int,
        help="Beam size khi giải mã (mặc định để Whisper tự quyết)",
    )
    parser.add_argument(
        "--best-of",
        type=int,
        help="Số lượng mẫu để chọn kết quả tốt nhất (mặc định để Whisper tự quyết)",
    )
    parser.add_argument(
        "--device",
        choices=("cpu", "gpu"),
        default="cpu",
        help="Chọn CPU hoặc GPU Nvidia để chạy suy luận (mặc định: cpu)",
    )
    return parser


def main(argv: Optional[Iterable[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    audio_path: Path = args.audio
    output_path: Path = args.output or _default_output_path(audio_path)

    options = TranscriptionOptions(
        model_name=args.model,
        language=args.language,
        task="translate" if args.translate else "transcribe",
        temperature=args.temperature,
        beam_size=args.beam_size,
        best_of=args.best_of,
        device="cuda" if args.device == "gpu" else "cpu",
    )

    output_file = transcribe_to_srt(audio_path, output_path, options=options)
    parser.exit(0, f"Đã xuất phụ đề ra {output_file}\n")


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
