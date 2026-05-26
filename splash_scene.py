# splash_scene.py
"""
Manim scene that renders the CPAuto splash animation to assets/splash.mp4.

Run once to (re)generate the video:
    python splash_scene.py

The rendered file is cached at assets/splash.mp4 and played by the GUI
on every subsequent launch without re-rendering.
"""

from __future__ import annotations

from pathlib import Path
import numpy as np

from manim import (
    BLACK, WHITE, DOWN, RIGHT, UP,
    Create, FadeIn, FadeOut,
    Flash, Scene, MoveAlongPath,
    Text, VGroup, VMobject, Write,
    config,
    linear,
)
from manim import ManimColor, rate_functions
from constants import SPLASH_ASSETS_DIR, SPLASH_VIDEO_PATH
# ---------------------------------------------------------------------------
# Output configuration
# ---------------------------------------------------------------------------

config.output_file     = "splash"
config.media_dir       = str(SPLASH_ASSETS_DIR)
config.video_dir       = str(SPLASH_ASSETS_DIR)
config.background_color = BLACK
config.pixel_height    = 720
config.pixel_width     = 1280
config.frame_rate      = 60

# ---------------------------------------------------------------------------
# Color palette
# ---------------------------------------------------------------------------

_CYAN_BLUE  = ManimColor("#00C8FF")
_TEAL       = ManimColor("#00E5CC")
_GOLD       = ManimColor("#FFD700")
_DARK_GRAY  = ManimColor("#1A1A2E")

# ---------------------------------------------------------------------------
# Scene
# ---------------------------------------------------------------------------

class CPAutoSplash(Scene):
    """
    Cinematic splash for CPAuto.

    Sequence:
      1. Background rectangle fades in (dark navy).
      2. "CP" assembles letter-by-letter from scattered glowing dots.
      3. "Auto" types in with a gradient colour sweep.
      4. A gold underline draws itself beneath the full title.
      5. Subtitle "Change Point Automation" fades in.
      6. Glow flash pulses outward from the title.
      7. Everything fades out to black.
    """

    def construct(self) -> None:
        # ---- Background ----
        from manim import Rectangle, ORIGIN
        bg = Rectangle(
            width=config.frame_width,
            height=config.frame_height,
            fill_color=_DARK_GRAY,
            fill_opacity=1,
            stroke_width=0,
        ).move_to(ORIGIN)
        self.add(bg)

        # ---- Compute layout positions without adding objects to the scene ----
        # Use a ghost VGroup to determine final positions only.
        _cp_ghost  = Text("CP",  font="Arial Black", font_size=120, color=_CYAN_BLUE, weight="BOLD")
        _a_ghost   = Text("A",   font="Arial Black", font_size=120, color=_TEAL,      weight="BOLD")
        _uto_ghost = Text("uto", font="Arial Black", font_size=120, color=_TEAL,      weight="BOLD")
        _layout    = VGroup(_cp_ghost, _a_ghost, _uto_ghost).arrange(RIGHT, buff=0.08)
        _layout.move_to(ORIGIN + UP * 0.4)

        cp_final  = _cp_ghost.get_center().copy()
        a_final   = _a_ghost.get_center().copy()
        uto_final = _uto_ghost.get_center().copy()
        title_left  = _layout.get_left().copy()
        title_right = _layout.get_right().copy()

        # ---- Real text objects, positioned at their final spots ----
        cp_text  = Text("CP",  font="Arial Black", font_size=120, color=_CYAN_BLUE, weight="BOLD").move_to(cp_final)
        a_text   = Text("A",   font="Arial Black", font_size=120, color=_TEAL,      weight="BOLD").move_to(a_final)
        uto_text = Text("uto", font="Arial Black", font_size=120, color=_TEAL,      weight="BOLD").move_to(uto_final)

        # ---- Underline ----
        from manim import Line
        underline = Line(
            start=title_left  + DOWN * 0.80,
            end=title_right   + DOWN * 0.80,
            color=_GOLD,
            stroke_width=4,
        )

        # ---- Subtitle ----
        subtitle = Text(
            "Change Point Automation",
            font="Arial",
            font_size=32,
            color=WHITE,
            slant="ITALIC",
        ).next_to(_layout, DOWN, buff=0.9)

        # ------------------------------------------------------------------------------------------- #
        # ---- Phase 1: "CP" writes itself with a glow flash ----
        self.play(
            Write(cp_text, run_time=0.8, rate_func=linear),
        )
        self.play(
            Flash(cp_text, color=_CYAN_BLUE, flash_radius=1.2, line_length=0.4, run_time=0.4),
        )

        # ------------------------------------------------------------------------------------------- #
        # ---- Phase 2: "A" drops vertically and bounces into place ----
        drop_start = np.array([
            a_final[0],
            a_final[1] + config.frame_height * 0.75,
            0
        ])

        a_text.move_to(drop_start)
        self.add(a_text)

        ground_y = a_final[1]

        bounce_height = [2.3, 1.6, 1.0, 0.5, 0.2]
        path_points = []
        current_y = drop_start[1]

        # ---------- Build vertical bounce path ----------
        for h in bounce_height:
            # Falling segment
            ys_down = np.linspace(current_y, ground_y, 40)
            xs_down = np.full_like(ys_down, a_final[0])

            down_pts = np.column_stack([xs_down, ys_down, np.zeros_like(xs_down)])
            path_points.extend(down_pts)

            # Rising segment (bounce back up)
            ys_up = np.linspace(ground_y, ground_y + h, 30)
            xs_up = np.full_like(ys_up, a_final[0])

            up_pts = np.column_stack([xs_up, ys_up, np.zeros_like(ys_up)])
            path_points.extend(up_pts)

            current_y = ground_y + h

        # Final fall to settle at the ground level
        ys_final = np.linspace(current_y, ground_y, 40)
        xs_final = np.full_like(ys_final, a_final[0])

        final_pts = np.column_stack([xs_final, ys_final, np.zeros_like(xs_final)])
        path_points.extend(final_pts)

        # Create the bounce path as a VMobject
        path = VMobject()
        path.set_points_smoothly(path_points)

        # ---------- Animate the "A" moving along the bounce path ----------
        self.play(
            MoveAlongPath(a_text, path), run_time = 1.8, rate_func = rate_functions.ease_in_sine
        )
        # Brief flash on landing
        self.play(
            Flash(a_text, color=_CYAN_BLUE, flash_radius=0.8, line_length=0.3, run_time=0.3),
        )

        # ------------------------------------------------------------------------------------------- #
        # ---- Phase 3: "uto" types in beside the settled "A" ----
        self.play(
            Write(uto_text, run_time=0.9, rate_func=rate_functions.ease_in_out_cubic),
        )

        # ------------------------------------------------------------------------------------------- #
        # ---- Phase 4: Gold underline draws itself ----
        self.play(
            Create(underline, run_time=0.5),
        )

        # ------------------------------------------------------------------------------------------- #
        # ---- Phase 5: Subtitle fades in ----
        self.play(
            FadeIn(subtitle, shift=UP * 0.2, run_time=0.6),
        )

        # ------------------------------------------------------------------------------------------- #
        # ---- Phase 6: Grand glow burst on the full title ----
        self.play(
            Flash(VGroup(cp_text, a_text, uto_text), color=_GOLD, flash_radius=2.5,
                  line_length=0.6, num_lines=20, run_time=0.6),
        )

        # ---- Hold ----
        self.wait(0.8)

        # ------------------------------------------------------------------------------------------- #
        # ---- Phase 7: Fade everything out ----
        self.play(
            FadeOut(VGroup(cp_text, a_text, uto_text, underline, subtitle), run_time=0.7),
        )
        self.wait(0.4)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import subprocess, sys
    subprocess.run(
        [
            sys.executable, "-m", "manim",
            __file__,
            "CPAutoSplash",
            "--format", "mp4",
            "-q", "h",          # high quality (1080p 60fps) — change to "m" for faster render
            "--disable_caching",
        ],
        check=True,
    )
    # Move rendered file to assets/splash_assets/splash.mp4
    import shutil
    candidates = list(Path("assets").rglob("CPAutoSplash.mp4"))
    if candidates:
        shutil.move(str(candidates[0]), str(SPLASH_VIDEO_PATH))
        print(f"[INFO] Splash video saved to: {SPLASH_VIDEO_PATH}")
    else:
        print("[WARN] Could not locate rendered CPAutoSplash.mp4 — check manim output.")
