from __future__ import annotations

from pathlib import Path

from .utils import FarFieldSource, build_cst_header, db_to_linear, format_linear, format_number, output_path_for


HEADER = "// >> Phi, Theta, Re(E_Theta), Im(E_Theta), Re(E_Phi), Im(E_Phi):"


def build_ffs_lines(source: FarFieldSource, mode: str, frequency_hz: float) -> list[str]:
    lines = build_cst_header(frequency_hz)
    lines.append("// >> Total #phi samples, total #theta samples")
    lines.append(f"{len(source.phi_angles)}   {len(source.theta_angles)}")
    lines.append("")
    lines.append(HEADER)
    for phi in source.phi_angles:
        for theta in source.theta_angles:
            key = (phi, theta)
            e_theta = db_to_linear(source.e_theta_db.get(key), mode)
            e_phi = db_to_linear(source.e_phi_db.get(key), mode)
            lines.append(
                " ".join(
                    [
                        format_number(phi),
                        format_number(theta),
                        format_linear(e_theta),
                        "0",
                        format_linear(e_phi),
                        "0",
                    ]
                )
            )
    return lines


def write_ffs(
    source: FarFieldSource,
    output_dir: str | Path,
    mode: str,
    frequency_hz: float,
) -> Path:
    path = output_path_for(source, output_dir, mode)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(build_ffs_lines(source, mode, frequency_hz)) + "\n", encoding="utf-8")
    return path
