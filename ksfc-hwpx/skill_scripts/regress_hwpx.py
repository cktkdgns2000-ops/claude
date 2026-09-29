#!/usr/bin/env python3
"""시험지 51건 회귀(hwpx): 마크업 → hwpx → 스키마 검증 → 쪽 배치 추정 → 원본(한글 저장본)의 실제 쪽 경계와 대조.

원본 hwpx 폴더(ORIG_HWPX, 기본 ./orig_hwpx)에 `<시험지 키>.hwpx`(예: kdx/2509_kdx_new_license_trend.hwpx)가 있으면
원본의 줄 배치 기록에서 실제 쪽별 첫 줄을 읽어 비교한다. 원본 .hwp는 corpus_tools/conv_all.py로 변환.
사용: python3 regress_hwpx.py [키 ...]
"""
import os
import re
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from hwpx_layout import Doc, norm  # noqa: E402

KEYS = [
    "2510_consortium",
    "2506_fsc_kdac",
    "2409_sto_trend",
    "2603_advisory_4th",
    "2604_fsc_investor_protection",
    "2607_mirae_collab",
    "2601_poc_explainer",
    "2605_prof_meeting",
    "2505_kdac_review",
    "kdx/2509_kdx_new_license_trend",
    "kdx/2510_kdx_investment_schedule",
    "kdx/2511_kdx_application_result",
    "kdx/2601_kdx_selection_result",
    "kdx/2602_kdx_briefing",
    "kdx/2603_kdx_sha_amendment",
    "kdx/2604_kdx_sha_further_amendment",
    "kdx/2607_kdx_rights_issue_issue",
    "kdx/2609_kdx_progress",
    "adv/2503_adv_operation_plan",
    "adv/2505_adv_formation_1st",
    "adv/2507_adv_2nd_plan",
    "adv/2511_adv_3rd",
    "adv/2607_adv_reorganization",
    "rsch/2504_rsch_proposal_prof",
    "rsch/2603_rsch_overview_prof",
    "rsch/2604_rsch_overview_internal",
    "rsch/2605_rsch_plan_internal",
    "rsch/2609_rsch_feedback_prof",
    "rsch/2609_rsch_progress_internal",
    "kscc/2410_kscc_mou_plan",
    "kscc/2507_kscc_joint_research_fsc",
    "kscc/2512_kscc_briefing_plan",
    "kscc/2512_kscc_progress_demo",
    "kscc/2609_kscc_poc_plan",
    "kdac/2504_kdac_meeting_result",
    "kdac/2509_kdac_trend_plan",
    "trend/2504_trend_custody_banks",
    "trend/2505_trend_sc_kscc_proposal",
    "trend/2506_trend_sc_legislation",
    "trend/2601_trend_vasp_phase2",
    "trend/2609_trend_sto_policy",
    "trend/2603_trend_sc_status",
    "trend/2608_trend_partnerships",
    "canton/2608_canton_sv_review",
    "canton/2609_canton_coop_plan",
    "equity/2606_equity_planning_review",
    "equity/2606_equity_feasibility",
    "misc/2603_misc_kofia_sc_proposal",
    "misc/2603_misc_council_plan",
    "misc/2604_misc_mirae_nda",
    "misc/2604_misc_council_result",
]   # assets/examples/<키>.md — 원본 재현 시험지 51건

ORIG = os.environ.get("ORIG_HWPX", os.path.join(ROOT, "orig_hwpx"))
OUT = os.environ.get("REGRESS_OUT", "/tmp/regress_hwpx")


def build(key):
    src = os.path.join(ROOT, "assets", "examples", key + ".md")
    dst = os.path.join(OUT, key + ".hwpx")
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    r = subprocess.run(["node", os.path.join(HERE, "build_hwpx.js"), src, dst], capture_output=True, text=True)
    if r.returncode:
        raise RuntimeError(r.stderr[-400:])
    return dst, r.stdout.strip().splitlines()[-1] if r.stdout.strip() else ""


def validate(path):
    """python-hwpx 스키마 검증. 통과 True, 실패 False, python-hwpx가 없으면 None"""
    r = subprocess.run([sys.executable, "-m", "hwpx.tools.validator", path], capture_output=True, text=True)
    if r.returncode != 0 and ("ModuleNotFoundError" in r.stderr or "ImportError" in r.stderr or "No module named" in r.stderr):
        return None
    return "All schema validations passed" in r.stdout


def leftover_markup(path):
    z = zipfile.ZipFile(path)
    s = z.read("Contents/section0.xml").decode()
    paras = ["".join(re.findall(r"<hp:t>(.*?)</hp:t>", p)) for p in re.findall(r"<hp:p .*?</hp:p>", s, re.S)]
    t = "\n".join(re.sub(r"^\s*\*\*\s", "", x) for x in paras)          # 줄 첫머리 '** '(둘째 각주 기호)는 정상
    return sorted(set(re.findall(r"\*\*[^*\s][^*]{0,40}?\*\*|__|\{\{|\}\}|\^\^|~~|==|!!|\{tight|\{small|\{right|\{narrow", t)))


def main(keys):
    tot = dict(docs=0, valid=0, pages=0, first=0, first_tot=0, pic=0)
    for key in keys:
        try:
            out, msg = build(key)
        except Exception as e:
            print(f"{key:40s} 빌드 실패: {e}")
            continue
        ok = validate(out)
        est = Doc(out).paginate()
        line = f"{key:40s} {'✓' if ok else ('-' if ok is None else '✗스키마')} {len(est)}쪽"
        orig = os.path.join(ORIG, key + ".hwpx")
        tot["docs"] += 1
        tot["valid"] += bool(ok)
        if os.path.exists(orig):
            act = Doc(orig, use_stored=True).actual_pages()
            same = sum(1 for a, b in zip(act, est) if norm(a) == norm(b))
            tot["pic"] += sum(1 for a, b in zip(act, est) if b == "(빈 쪽)" and norm(a) != norm(b))   # 원고가 원본 쪽을 그림으로 넣은 쪽
            tot["pages"] += len(act) == len(est)
            tot["first"] += same
            tot["first_tot"] += len(act)
            line += f" / 원본 {len(act)}쪽, 쪽 첫 줄 {same}/{len(act)}"
            if len(act) != len(est) or same < len(act):
                diff = [(i + 1, a[:10], b[:10]) for i, (a, b) in enumerate(zip(act, est)) if norm(a) != norm(b)][:2]
                line += f"  {diff}"
        lo = leftover_markup(out)
        if lo:
            line += f"  [잔여 마크업 {lo}]"
        print(line, "|", msg.split("|", 1)[-1].strip() if "|" in msg else "")
    print(f"\n문서 {tot['docs']} · 스키마 통과 {tot['valid']} · 쪽수 일치 {tot['pages']} · 쪽 첫 줄 일치 {tot['first']}/{tot['first_tot']}"
          f" (그림으로 넣은 쪽 {tot['pic']}개 제외 시 {tot['first']}/{tot['first_tot'] - tot['pic']})")


if __name__ == "__main__":
    main(sys.argv[1:] or KEYS)
