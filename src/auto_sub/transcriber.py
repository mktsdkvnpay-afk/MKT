"""Core transcription logic for converting audio to SRT subtitles."""

from __future__ import annotations

from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path
from typing import Iterable, List, Optional, Sequence


@dataclass(frozen=True)
class TranscriptionOptions:
    """Configuration for the transcription process."""

    model_name: str = "base"
    language: Optional[str] = None
    task: str = "transcribe"  # "transcribe" or "translate"
    temperature: float = 0.0
    beam_size: Optional[int] = None
    best_of: Optional[int] = None
    device: str = "cpu"


def _format_timestamp(seconds: float) -> str:
    milliseconds = max(int(round(seconds * 1000)), 0)
    hours, milliseconds = divmod(milliseconds, 3_600_000)
    minutes, milliseconds = divmod(milliseconds, 60_000)
    secs, milliseconds = divmod(milliseconds, 1_000)
    return f"{hours:02}:{minutes:02}:{secs:02},{milliseconds:03}"


def _coerce_segments(raw_segments: Sequence[dict]) -> List[dict]:
    segments: List[dict] = []
    for segment in raw_segments:
        text = segment.get("text", "").strip()
        if not text:
            continue
        segments.append(
            {
                "start": float(segment.get("start", 0.0)),
                "end": float(segment.get("end", 0.0)),
                "text": text,
            }
        )
    return segments


def _normalise_device(device: str) -> str:
    device = (device or "cpu").lower()
    if device in {"gpu", "cuda"}:
        return "cuda"
    if device == "cpu":
        return "cpu"
    raise ValueError(f"Thiết bị không được hỗ trợ: {device}")


def _ensure_device_available(device: str) -> str:
    normalised = _normalise_device(device)
    if normalised == "cuda":
        try:
            import torch  # type: ignore
        except ImportError as exc:  # pragma: no cover - defensive import guard
            raise ImportError(
                "Cần cài đặt PyTorch hỗ trợ CUDA để chạy Whisper trên GPU Nvidia."
            ) from exc
        if not torch.cuda.is_available():
            raise RuntimeError(
                "Không tìm thấy GPU Nvidia khả dụng. Kiểm tra driver và CUDA runtime."
            )
    return normalised


@lru_cache(maxsize=4)
def _load_model(model_name: str, device: str):
    try:
        import whisper  # type: ignore
    except ImportError as exc:  # pragma: no cover - defensive import guard
        raise ImportError(
            "Cần cài đặt gói 'openai-whisper' để sử dụng chức năng tạo phụ đề."
        ) from exc

    return whisper.load_model(model_name, device=device)


def segments_to_srt(segments: Iterable[dict]) -> str:
    """Convert iterable of Whisper segments to SRT formatted subtitles."""

    lines: List[str] = []
    for index, segment in enumerate(segments, start=1):
        start_ts = _format_timestamp(segment["start"])
        end_ts = _format_timestamp(segment["end"])
        text = segment["text"].replace(" --> ", " → ")
        lines.extend([str(index), f"{start_ts} --> {end_ts}", text, ""])
    return "\n".join(lines).strip() + "\n"


def transcribe(audio_path: Path, *, options: TranscriptionOptions | None = None) -> List[dict]:
    """Transcribe ``audio_path`` using Whisper and return normalised segments."""

    audio_path = Path(audio_path)
    if options is None:
        options = TranscriptionOptions()

    if not audio_path.exists():
        raise FileNotFoundError(f"Audio file not found: {audio_path}")

    device = _ensure_device_available(options.device)

    inference_args = {
        "temperature": options.temperature,
    }
    if options.language:
        inference_args["language"] = options.language
    if options.task:
        inference_args["task"] = options.task
    if options.beam_size is not None:
        inference_args["beam_size"] = options.beam_size
    if options.best_of is not None:
        inference_args["best_of"] = options.best_of

    model = _load_model(options.model_name, device)
    result = model.transcribe(str(audio_path), **inference_args)
    return _coerce_segments(result.get("segments", []))


def transcribe_to_srt(audio_path: Path, output_path: Path, *, options: TranscriptionOptions | None = None) -> Path:
    """Transcribe audio and write the SRT subtitle file."""

    segments = transcribe(audio_path, options=options)
    srt_text = segments_to_srt(segments)

    output_path = Path(output_path)
    output_path.write_text(srt_text, encoding="utf-8")
    return output_path
