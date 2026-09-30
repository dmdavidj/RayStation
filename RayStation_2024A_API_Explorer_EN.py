# -*- coding: utf-8 -*-
u"""
RayStation 2024A - Scripting API Explorer (GUI) [English]
=========================================================

KO: 실행 중인 RayStation 의 스크립팅 API 를 직접 탐색해서 객체 타입별 속성/메서드,
    파라미터, docstring, 자동 해설, 버전·빌드 유형(Clinical/Research/Evaluation 등)을
    Markdown(+ 버전 비교용 JSON)으로 저장합니다.
EN: Introspects the scripting API of the running RayStation and saves every object
    type's properties/methods, parameters, docstrings, automatic notes and the detected
    version / build type (Clinical/Research/Evaluation ...) as Markdown (+ JSON for
    version comparison).

Usage / 사용법
  1. Open a patient, case, plan and beam set in RayStation.
     (환자·Case·Plan·BeamSet 을 열어 둡니다. 데이터가 풍부할수록 더 많은 타입이 발견됩니다.)
  2. Run this file from Scripting > Script Management.
  3. Choose "Save As..." (or type a full path), then press Run.

Safety / 안전성
  - Methods of RayStation objects are NEVER called; only property values are read.
  - Nothing is modified or saved in RayStation.
  - Window titles are not stored (they can contain patient names); only version /
    build keywords extracted from them are reported.
  - String/date values are masked by default ("Mask text values" option).

Compatibility: CPython 3.x (pythonnet) and IronPython 2.7. GUI uses System.Windows.Forms.
"""
from __future__ import print_function, unicode_literals

import sys
import os
import re
import json
import time
import codecs
import hashlib
import datetime
import inspect
import traceback

# ---------------------------------------------------------------------------
# Per-file settings (these are the only lines that differ between the 4 files)
# ---------------------------------------------------------------------------
SCRIPT_TARGET_VERSION = "2024A"
LANG = "en"            # "ko" = Korean report/UI, "en" = English report/UI
SCRIPT_TITLE = "RayStation 2024A API Explorer (English)"

ROOT_KEYS = [
    ("Patient", True),
    ("Case", True),
    ("Examination", True),
    ("Plan", True),
    ("BeamSet", True),
    ("PatientDB", True),
    ("MachineDB", True),
    ("ClinicDB", True),
    ("ui", False),          # the UI tree is very large - enable only when needed
]

DEFAULT_SKIP = [
    "PixelData", "DoseValues", "Contours", "Vertices", "Triangles",
    "Normals", "Indices", "VoxelData",
]

DEFAULT_MAX_DEPTH = 10
DEFAULT_MAX_TYPES = 2000
DEFAULT_SAMPLES = 3

# ---------------------------------------------------------------------------
# Python 2/3 compatibility
# ---------------------------------------------------------------------------
PY3 = sys.version_info[0] >= 3
if PY3:
    string_types = (str, bytes)
    integer_types = (int,)
else:  # IronPython 2.7
    string_types = (basestring,)  # noqa: F821
    integer_types = (int, long)  # noqa: F821
NUMBER_TYPES = integer_types + (float, complex)

try:
    import connect  # RayStation scripting module
except Exception:
    connect = None

NET_NOISE = set([
    "Equals", "GetHashCode", "GetType", "ToString", "ReferenceEquals",
    "MemberwiseClone", "Finalize", "Overloads",
])
GENERIC_CLASS_HINTS = ("ScriptObject", "Proxy", "Wrapper", "Dynamic", "Expando")
DEPRECATED_RE = re.compile(r"deprecat|obsolete|will be removed|no longer supported|폐기", re.I)


# ---------------------------------------------------------------------------
# Strings (ko, en)
# ---------------------------------------------------------------------------
S = {
    # ---- GUI
    "gui_title": ("RayStation %s 스크립팅 API 구조 추출 → Markdown",
                  "RayStation %s Scripting API Extractor → Markdown"),
    "gui_sub": ("메서드는 호출하지 않고 속성 값만 읽습니다. 환자·Plan·BeamSet 을 열어 둔 상태에서 실행하세요.",
                "Methods are never called; only property values are read. Open a patient, plan and beam set first."),
    "gui_detected": ("감지된 환경", "Detected"),
    "gui_label": ("버전 라벨", "Version label"),
    "gui_label_hint": ("문서 제목·파일명에 사용 (자동 감지값, 수정 가능)",
                       "Used in the title and file name (auto-detected, editable)"),
    "gui_save": ("저장 파일", "Save to"),
    "gui_save_hint": ("경로를 직접 입력하거나 [다른 이름으로...] 로 지정하세요. 폴더만 입력하면 기본 파일명이 붙습니다.",
                      "Type a full path or use [Save As...]. If you enter only a folder, a default file name is added."),
    "gui_browse": ("다른 이름으로...", "Save As..."),
    "gui_roots": ("탐색 루트 (get_current)", "Roots (get_current)"),
    "gui_depth": ("최대 탐색 깊이", "Max depth"),
    "gui_types": ("최대 타입 수", "Max types"),
    "gui_samples": ("컬렉션 샘플 수", "Collection samples"),
    "gui_samples_hint": ("컬렉션에서 여러 요소를 살펴보면 다른 형상·빔 타입도 발견합니다.",
                         "Inspecting several elements finds more geometry / beam types."),
    "gui_json": ("JSON 도 함께 저장 (버전 비교용)", "Also save JSON (for version comparison)"),
    "gui_mask": ("문자열·날짜 값 가리기 (환자 정보 보호)", "Mask text / date values (protects patient data)"),
    "gui_skip": ("값을 읽지 않을 대용량 속성 (쉼표 구분)", "Heavy properties to skip (comma separated)"),
    "gui_run": ("추출 실행", "Run"),
    "gui_stop": ("중지", "Stop"),
    "gui_close": ("닫기", "Close"),
    "gui_idle": ("대기 중", "Ready"),
    "gui_progress": ("탐색 %d / %d (발견 타입 %d) : %s", "Exploring %d / %d (types found %d): %s"),
    "gui_done": ("완료 (%.1f초)", "Done (%.1f s)"),
    "gui_error": ("오류 발생 - 로그 확인", "Error - see log"),
    "gui_dlg_save": ("API 문서를 저장할 위치와 파일 이름을 지정하세요",
                     "Choose where to save the API document"),
    "gui_need_path": ("저장 파일 경로를 입력하세요.", "Enter a file path to save to."),
    "gui_need_root": ("탐색 루트를 하나 이상 선택하세요.", "Select at least one root."),
    "gui_overwrite": ("같은 이름의 파일이 이미 있습니다. 덮어쓸까요?\n\n%s",
                      "A file with this name already exists. Overwrite?\n\n%s"),
    "gui_saved": ("저장 완료%s\n\n%s\n\n타입 %d · 메서드 %d · 속성 %d\n\n파일 위치를 열까요?",
                  "Saved%s\n\n%s\n\nTypes %d · Methods %d · Properties %d\n\nOpen the file location?"),
    "gui_partial": (" (중지됨 - 일부 결과)", " (stopped - partial result)"),
    "gui_fail": ("추출 중 오류가 발생했습니다.\n\n%s", "An error occurred during extraction.\n\n%s"),
    "gui_bad_path": ("저장 경로를 사용할 수 없습니다.\n\n%s", "The save path cannot be used.\n\n%s"),
    # ---- log
    "log_start": ("RayStation %s API 추출 시작", "Starting RayStation %s API extraction"),
    "log_env": ("환경 감지: %s", "Environment: %s"),
    "log_roots": ("루트 객체 취득 중...", "Getting root objects..."),
    "log_globals": ("connect 전역 함수 수집 중...", "Collecting connect module functions..."),
    "log_globals_n": ("  전역 함수 %d개", "  %d global functions"),
    "log_walk": ("객체 트리 탐색 중... (메서드는 호출하지 않습니다)",
                 "Walking the object tree... (methods are not called)"),
    "log_md": ("Markdown 저장: %s", "Markdown saved: %s"),
    "log_json": ("JSON 저장: %s", "JSON saved: %s"),
    "log_done": ("완료: 타입 %d · 메서드 %d · 속성 %d", "Done: types %d · methods %d · properties %d"),
    "log_stop": ("중지 요청됨 - 지금까지의 결과로 문서를 저장합니다.",
                 "Stop requested - the document will be saved with the results so far."),
    "log_fail_type": ("  ! %s 탐색 실패: %s", "  ! failed to explore %s: %s"),
    "log_err": ("오류: %s", "Error: %s"),
    # ---- walker
    "p_skipped_t": ("(값 읽지 않음)", "(not read)"),
    "p_skipped_n": ("대용량 데이터 추정 - 건너뛰도록 지정됨", "Assumed heavy data - configured to skip"),
    "p_error_t": ("(현재 상태에서 접근 불가)", "(not accessible in current state)"),
    "p_fail_t": ("(분석 실패)", "(analysis failed)"),
    "p_empty_t": ("Collection (비어 있음)", "Collection (empty)"),
    "p_count": ("%d개", "%d items"),
    "p_masked": ("\"***\" (%d자)", "\"***\" (%d chars)"),
    "p_masked2": ("*** (가림)", "*** (masked)"),
    "elem": ("%s 요소", "%s element"),
    "no_connect": ("connect 모듈을 불러올 수 없음 (RayStation 밖에서 실행?)",
                   "Cannot import connect (running outside RayStation?)"),
    "none_ret": ("None 반환", "returned None"),
    # ---- status
    "st_explored": ("탐색 완료", "explored"),
    "st_partial": ("일부 탐색(중지됨)", "partially explored (stopped)"),
    "st_depth_limit": ("⛔ 깊이 제한으로 미탐색", "⛔ not explored (depth limit)"),
    "st_type_limit": ("⛔ 타입 수 제한으로 미탐색", "⛔ not explored (type limit)"),
    "st_cancelled": ("⛔ 중지로 미탐색", "⛔ not explored (stopped)"),
    "st_error": ("❗ 탐색 오류", "❗ exploration error"),
    "st_pending": ("미탐색", "not explored"),
    # ---- build types
    "bt_clinical": ("Clinical (임상용)", "Clinical"),
    "bt_clinical_assumed": ("Clinical 추정 (Research/Evaluation 등 표식 없음)",
                            "Presumably Clinical (no Research/Evaluation marker found)"),
    "bt_research": ("Research (연구용 · 비임상)", "Research (non-clinical)"),
    "bt_evaluation": ("Evaluation (평가판)", "Evaluation"),
    "bt_nonclinical": ("Non-clinical (비임상용 표기)", "Non-clinical (marked not for clinical use)"),
    "bt_validation": ("Validation / Pre-release (검증·사전 배포판)", "Validation / Pre-release"),
    "bt_training": ("Training / Education (교육용)", "Training / Education"),
    "bt_unknown": ("알 수 없음 (RayStation 정보를 찾지 못함)", "Unknown (no RayStation information found)"),
    # ---- report
    "r_title": ("RayStation %s Scripting API 레퍼런스 (자동 추출)",
                "RayStation %s Scripting API Reference (auto-extracted)"),
    "r_intro": ("실행 중인 RayStation 에서 `dir()` 과 docstring 으로 **직접 추출**한 문서입니다. "
                "**자동 해설**은 메서드/속성 이름으로 추정한 설명이므로, 정확한 의미·제약은 "
                "**공식 설명(docstring)** 과 RayStation Scripting API 문서를 우선하세요.",
                "Extracted **directly** from the running RayStation with `dir()` and docstrings. "
                "**Auto notes** are guesses based on member names; for exact meaning and constraints, "
                "rely on the **official description (docstring)** and RaySearch's scripting API documentation."),
    "r_item": ("항목", "Item"),
    "r_value": ("값", "Value"),
    "r_created": ("📄 파일 생성 시점", "📄 File created"),
    "r_label": ("버전 라벨", "Version label"),
    "r_target": ("스크립트 대상 버전", "Script target version"),
    "r_detver": ("감지된 RayStation 버전", "Detected RayStation version"),
    "r_match": ("✅ 대상 버전과 일치", "✅ matches target"),
    "r_mismatch": ("⚠️ 대상 버전(%s)과 다름 — 다른 버전용 스크립트를 실행했는지 확인하세요",
                   "⚠️ differs from target (%s) — check you ran the right script"),
    "r_nover": ("감지 못함", "not detected"),
    "r_prodver": ("제품 버전 (ProductVersion)", "Product version"),
    "r_buildtype": ("🏷️ 빌드 유형", "🏷️ Build type"),
    "r_basis": ("판단 근거", "Basis"),
    "r_exe": ("RayStation 실행 파일", "RayStation executable"),
    "r_exe_mtime": ("실행 파일 수정일 (설치/빌드 시점 참고)", "Executable modified (install/build hint)"),
    "r_python": ("Python", "Python"),
    "r_script": ("스크립트", "Script"),
    "r_settings": ("탐색 설정", "Settings"),
    "r_settings_v": ("깊이 %d · 최대 타입 %d · 컬렉션 샘플 %d · 값 가리기 %s",
                     "depth %d · max types %d · collection samples %d · masking %s"),
    "r_on": ("켬", "on"),
    "r_off": ("끔", "off"),
    "r_elapsed": ("소요 시간", "Elapsed"),
    "r_sec": ("%.1f 초", "%.1f s"),
    "r_env_h": ("실행 환경 / 빌드 정보 근거", "Environment / build evidence"),
    "r_env_note": ("버전·빌드 유형은 아래 정보에서 자동 판단했습니다. 창 제목은 환자 이름이 포함될 수 있어 "
                   "원문을 저장하지 않고 버전·빌드 키워드만 기록합니다.",
                   "Version and build type were inferred from the information below. Window titles can contain "
                   "patient names, so only version/build keywords extracted from them are recorded."),
    "r_source": ("출처", "Source"),
    "r_roots_h": ("루트 객체 취득 결과", "Root objects"),
    "r_root": ("루트", "Root"),
    "r_result": ("결과", "Result"),
    "r_note": ("비고", "Note"),
    "r_ok": ("✅ 취득", "✅ obtained"),
    "r_fail": ("❌ 실패", "❌ failed"),
    "r_summary": ("요약", "Summary"),
    "r_kind": ("구분", "Kind"),
    "r_count": ("개수", "Count"),
    "r_n_types": ("발견된 객체 타입", "Object types found"),
    "r_n_methods": ("메서드 (객체)", "Methods (objects)"),
    "r_n_globals": ("전역 함수 (connect)", "Global functions (connect)"),
    "r_n_props": ("속성", "Properties"),
    "r_n_dep": ("⚠️ Deprecated 표시 메서드", "⚠️ Methods marked deprecated"),
    "r_n_err": ("🚫 현재 상태에서 접근 불가 속성", "🚫 Properties not accessible now"),
    "r_n_empty": ("∅ 비어 있는 컬렉션 (하위 타입 미확인)", "∅ Empty collections (child type unknown)"),
    "r_n_unexp": ("⛔ 미탐색 타입", "⛔ Unexplored types"),
    "r_legend": ("**기호**: ⚠️ Deprecated · 🚫 현재 상태에서 접근 불가(예: 선량 미계산, 해당 객체 없음) · "
                 "⏭ 값 읽기 생략 · ∅ 빈 컬렉션 · ⛔ 미탐색",
                 "**Legend**: ⚠️ deprecated · 🚫 not accessible in current state (e.g. dose not computed) · "
                 "⏭ value not read · ∅ empty collection · ⛔ not explored"),
    "r_toc": ("목차", "Contents"),
    "r_idx_h": ("기능 분류 인덱스", "Feature index"),
    "r_idx_note": ("메서드 이름의 동사로 분류했습니다. 이 버전에서 **무엇을 할 수 있는지** 한눈에 보는 용도입니다.",
                   "Methods grouped by the verb in their name, to see at a glance **what this version can do**."),
    "r_cat": ("분류", "Category"),
    "r_n_m": ("메서드 수", "Methods"),
    "r_warn_h": ("주의 목록", "Watch list"),
    "r_dep_h": ("⚠️ Deprecated 로 표시된 메서드", "⚠️ Methods marked deprecated"),
    "r_dep_none": ("없음 (docstring 에 deprecated/obsolete 표기 없음)", "None (no deprecated/obsolete in docstrings)"),
    "r_err_h": ("🚫 현재 상태에서 접근 불가 속성", "🚫 Properties not accessible in the current state"),
    "r_err_note": ("현재 열린 환자/계획 상태 때문에 읽을 수 없었던 속성입니다. **기능이 없다는 뜻은 아닙니다** "
                   "(예: 선량이 계산되지 않았거나, 해당 모달리티가 아닌 경우).",
                   "These could not be read because of the currently open patient/plan. **This does not mean the "
                   "feature is missing** (e.g. dose not computed, different modality)."),
    "r_where": ("위치", "Where"),
    "r_prop": ("속성", "Property"),
    "r_error": ("오류", "Error"),
    "r_none": ("없음", "None"),
    "r_empty_h": ("∅ 비어 있는 컬렉션", "∅ Empty collections"),
    "r_empty_note": ("요소가 없어 하위 타입을 탐색하지 못했습니다. 해당 데이터를 가진 환자로 다시 실행하면 문서가 더 완전해집니다.",
                     "No elements, so the child type could not be explored. Re-run with a patient that has this data."),
    "r_unexp_h": ("⛔ 미탐색 타입", "⛔ Unexplored types"),
    "r_glob_h": ("전역 함수 (connect 모듈)", "Global functions (connect module)"),
    "r_glob_note": ("스크립트 맨 위의 `from connect import *` 로 사용하는 함수/클래스입니다.",
                    "Functions/classes available through `from connect import *`."),
    "r_glob_fail": ("connect 모듈을 읽지 못했습니다.", "Could not read the connect module."),
    "r_glob_other": ("기타 전역 항목", "Other globals"),
    "r_name": ("이름", "Name"),
    "r_type": ("타입", "Type"),
    "r_valnote": ("값/비고", "Value / note"),
    "r_types_h": ("타입별 상세", "Types in detail"),
    "r_n_mp": ("메서드 %d · 속성 %d", "methods %d · properties %d"),
    "r_paths": ("접근 경로", "Access paths"),
    "r_code": ("코드 예", "Code"),
    "r_internal": ("내부 타입명", "Internal type"),
    "r_state": ("상태", "Status"),
    "r_props_h": ("속성", "Properties"),
    "r_props_cols": ("| 속성 | 타입 | 현재 값(예) | 자동 해설 |", "| Property | Type | Current value (example) | Auto note |"),
    "r_methods_h": ("메서드", "Methods"),
    "r_methods_cols": ("| 메서드 | 분류 | 자동 해설 |", "| Method | Category | Auto note |"),
    "r_sig": ("호출 형태", "Signature"),
    "r_nosig": ("(…)  ← 시그니처 정보 없음", "(…)  ← no signature info"),
    "r_auto": ("자동 해설", "Auto note"),
    "r_official": ("공식 설명", "Official description"),
    "r_nodoc": ("_(docstring 없음 — RayStation Scripting API 문서 참고)_",
                "_(no docstring — see the RayStation scripting API documentation)_"),
    "r_returns": ("반환", "Returns"),
    "r_anerr": ("분석 오류", "Analysis error"),
    "r_param_cols": ("| 파라미터 | 타입 | 기본값 | 필수/선택 | 설명 |",
                     "| Parameter | Type | Default | Required | Description |"),
    "r_opt": ("선택", "optional"),
    "r_req": ("필수/미상", "required / unknown"),
    "r_val": ("값", "value"),
    "r_fulldoc": ("전체 docstring", "Full docstring"),
    "r_target_obj": ("대상", "the object"),
    "cat_other": ("기타", "Other"),
}


def tr(key, *args):
    pair = S[key]
    s = pair[0] if LANG == "ko" else pair[1]
    return s % args if args else s


# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------
def to_text(v):
    try:
        if PY3:
            if isinstance(v, bytes):
                return v.decode("utf-8", "replace")
            return str(v)
        if isinstance(v, unicode):  # noqa: F821
            return v
        return str(v).decode("utf-8", "replace")
    except Exception:
        try:
            return repr(v)
        except Exception:
            return "<?>"


def short(v, n=70):
    t = to_text(v).replace("\r", " ").replace("\n", " ")
    t = " ".join(t.split())
    return t if len(t) <= n else t[:n - 1] + "…"


def err_text(e):
    t = to_text(e).strip().splitlines()
    t = t[0] if t else type(e).__name__
    return short("%s: %s" % (type(e).__name__, t), 220)


def md_escape(s):
    return to_text(s).replace("|", "\\|").replace("\r", "").replace("\n", "<br>")


def safe_dir(obj):
    try:
        return list(dir(obj))
    except Exception:
        return []


def api_names(obj):
    return [n for n in safe_dir(obj)
            if n and not n.startswith("_") and n not in NET_NOISE]


def has_api_members(obj):
    for n in api_names(obj):
        if n[:1].isupper():
            return True
    return False


def dotnet_type_name(obj):
    try:
        t = obj.GetType()
        n = to_text(t.FullName)
        if n and n != "None":
            return n
    except Exception:
        pass
    return None


def display_type_name(obj):
    n = dotnet_type_name(obj)
    if n and not n.startswith("System.") and not n.startswith("Python."):
        base = n.split(".")[-1]
        if not any(h in base for h in GENERIC_CLASS_HINTS):
            return base
    try:
        cls = type(obj).__name__
        mod = getattr(type(obj), "__module__", "") or ""
        if cls and mod not in ("builtins", "__builtin__") and \
                not any(h in cls for h in GENERIC_CLASS_HINTS):
            return cls
    except Exception:
        pass
    return None


def sig_key(names):
    s = ",".join(sorted(names)).encode("utf-8")
    return "t" + hashlib.md5(s).hexdigest()[:10]


def normalize_path(expr):
    return re.sub(r"\[\d+\]", "[]", expr)


def root_code(expr):
    m = re.match(r"^([A-Za-z_]\w*)(.*)$", expr)
    if not m:
        return expr
    root, rest = m.group(1), m.group(2)
    if root == "connect":
        return "connect" + rest
    return 'get_current("%s")%s' % (root, rest)


def now_with_offset():
    now = datetime.datetime.now()
    try:
        off = now - datetime.datetime.utcnow()
        mins = int(round(off.total_seconds() / 60.0))
    except Exception:
        mins = 0
    sign = "+" if mins >= 0 else "-"
    mins = abs(mins)
    return now, "%s (UTC%s%02d:%02d)" % (now.strftime("%Y-%m-%d %H:%M:%S"), sign, mins // 60, mins % 60)


# ---------------------------------------------------------------------------
# Automatic notes (name based)
# ---------------------------------------------------------------------------
VERB_KO = {
    "Create": "생성", "Add": "추가", "Make": "생성", "Generate": "생성", "New": "새로 만들기",
    "Append": "추가", "Insert": "삽입", "Delete": "삭제", "Remove": "제거", "Clear": "비우기",
    "Unload": "언로드", "Discard": "폐기", "Set": "설정", "Get": "조회", "Update": "갱신",
    "Edit": "편집", "Change": "변경", "Modify": "수정", "Assign": "할당", "Rename": "이름 변경",
    "Reset": "초기화", "Compute": "계산", "Calculate": "계산", "Recompute": "재계산",
    "Run": "실행", "Execute": "실행", "Optimize": "최적화", "Reoptimize": "재최적화",
    "Evaluate": "평가", "Estimate": "추정", "Simulate": "시뮬레이션", "Import": "가져오기",
    "Export": "내보내기", "Load": "불러오기", "Save": "저장", "Store": "저장", "Send": "전송",
    "Print": "출력", "Open": "열기", "Close": "닫기", "Copy": "복사", "Duplicate": "복제",
    "Convert": "변환", "Transform": "변환", "Merge": "병합", "Split": "분할", "Mirror": "대칭 복사",
    "Apply": "적용", "Select": "선택", "Move": "이동", "Rotate": "회전", "Translate": "평행 이동",
    "Scale": "스케일 조정", "Rescale": "재스케일", "Normalize": "정규화", "Toggle": "토글",
    "Approve": "승인", "Unapprove": "승인 해제", "Lock": "잠금", "Unlock": "잠금 해제",
    "Accept": "수락", "Reject": "거부", "Cancel": "취소", "Query": "검색", "Find": "찾기",
    "Search": "검색", "List": "목록 조회", "Is": "여부 확인", "Has": "보유 여부 확인",
    "Can": "가능 여부 확인", "Check": "검사", "Validate": "검증", "Verify": "확인",
    "Register": "등록/정합", "Deform": "변형", "Adapt": "적응 재계획", "Resample": "리샘플",
    "Simplify": "단순화", "Smooth": "스무딩", "Expand": "확장", "Contract": "축소",
    "Interpolate": "보간", "Crop": "자르기", "Fill": "채우기", "Threshold": "임계값 처리",
    "Grow": "영역 확장", "Fit": "맞춤", "Show": "표시", "Hide": "숨기기", "Refresh": "새로고침",
    "Sort": "정렬", "Replace": "교체", "Round": "반올림", "Snap": "맞춤 정렬", "Link": "연결",
    "Unlink": "연결 해제", "Enable": "활성화", "Disable": "비활성화", "Sample": "샘플링",
    "Scan": "스캔", "Predict": "예측", "Train": "학습", "Segment": "분할(세그멘테이션)",
}

TERM_KO = {
    "Roi": "ROI(관심 구조물)", "Rois": "ROI 목록", "Poi": "POI(관심 점)", "Pois": "POI 목록",
    "RegionsOfInterest": "ROI 목록", "PointsOfInterest": "POI 목록",
    "BeamSet": "빔셋", "BeamSets": "빔셋 목록", "Beam": "빔", "Beams": "빔 목록",
    "Plan": "치료계획", "Plans": "치료계획 목록", "TreatmentPlan": "치료계획",
    "TreatmentPlans": "치료계획 목록", "Dose": "선량", "Doses": "선량", "Dvh": "DVH",
    "Optimization": "최적화", "Optimizations": "최적화 목록", "Isocenter": "아이소센터",
    "Examination": "영상 검사(CT/MR 등)", "Examinations": "영상 검사 목록",
    "Registration": "영상 정합", "Registrations": "영상 정합 목록", "Structure": "구조",
    "StructureSet": "구조셋", "StructureSets": "구조셋 목록", "Geometry": "형상",
    "Geometries": "형상 목록", "Contour": "윤곽", "Contours": "윤곽", "Segment": "세그먼트",
    "Segments": "세그먼트 목록", "Mlc": "MLC", "Jaw": "Jaw", "Couch": "카우치",
    "Gantry": "갠트리", "Collimator": "콜리메이터", "Arc": "아크", "Fraction": "분할(Fraction)",
    "Fractions": "분할 수", "Prescription": "처방", "ClinicalGoal": "임상 목표(Clinical Goal)",
    "ClinicalGoals": "임상 목표 목록", "Objective": "목적함수", "Objectives": "목적함수 목록",
    "Constraint": "제약조건", "Constraints": "제약조건 목록", "Material": "물질",
    "Density": "밀도", "Margin": "마진", "Algebra": "불리언 연산", "Template": "템플릿",
    "Machine": "치료기", "Patient": "환자", "Case": "케이스", "Cases": "케이스 목록",
    "Series": "시리즈", "Image": "영상", "Images": "영상", "Stack": "스택", "Grid": "그리드",
    "Report": "리포트", "Dicom": "DICOM", "Evaluation": "평가", "Robust": "강건(Robust)",
    "Robustness": "강건성", "Qa": "QA", "Deformable": "변형(Deformable)", "Rigid": "강체",
    "Mbs": "모델 기반 분할(MBS)", "Segmentation": "세그멘테이션", "Bolus": "볼루스",
    "Wedge": "웨지", "Energy": "에너지", "Spot": "스팟", "Spots": "스팟 목록", "Layer": "레이어",
    "Proton": "양성자", "Electron": "전자", "Photon": "광자", "Brachy": "근접치료",
    "Catheter": "카테터", "Applicator": "어플리케이터", "Weight": "가중치", "Mu": "MU",
    "Clinic": "클리닉", "Db": "DB", "Name": "이름", "Type": "유형", "Color": "색상",
    "Comment": "코멘트", "Description": "설명", "Id": "ID", "Position": "위치",
    "Point": "점", "Points": "점 목록", "Volume": "부피", "Center": "중심", "Box": "박스",
    "Sphere": "구", "Cylinder": "원기둥", "Mesh": "메시", "Voxel": "복셀", "Frame": "프레임",
    "FrameOfReference": "좌표계(Frame of Reference)", "Tissue": "조직", "Setup": "셋업",
    "Treatment": "치료", "Delivery": "전달", "Technique": "기법", "Algorithm": "알고리즘",
    "Settings": "설정값", "Setting": "설정값", "Parameters": "파라미터", "Function": "함수",
    "Functions": "함수 목록", "Value": "값", "Values": "값", "Level": "레벨", "Isodose": "등선량",
    "Approval": "승인", "Status": "상태", "Uid": "UID", "Sop": "SOP", "Study": "Study",
    "Collision": "충돌", "Field": "조사야", "Aperture": "조리개(Aperture)", "Block": "블록",
    "Leaf": "리프", "Leaves": "리프", "Control": "제어", "Plane": "평면", "Slice": "슬라이스",
    "Hu": "HU", "Ct": "CT", "Mr": "MR", "Pet": "PET", "Cbct": "CBCT", "Ui": "UI",
    "Window": "창", "Workspace": "작업 공간", "Module": "모듈", "Tab": "탭", "Button": "버튼",
    "Mean": "평균", "Max": "최대", "Min": "최소", "Relative": "상대", "Absolute": "절대",
    "Model": "모델", "PatientModel": "환자 모델", "Target": "타깃", "External": "외곽(External)",
    "Organ": "장기", "Oar": "OAR", "Ptv": "PTV", "Ctv": "CTV", "Gtv": "GTV",
    "Statistics": "통계", "Statistic": "통계", "Scenario": "시나리오", "Scenarios": "시나리오 목록",
    "Uncertainty": "불확도", "Override": "오버라이드", "Script": "스크립트",
}

TERM_EN = {
    "Roi": "ROI", "Rois": "ROIs", "Poi": "POI", "Pois": "POIs", "Dvh": "DVH", "Mlc": "MLC",
    "Dicom": "DICOM", "Qa": "QA", "Mu": "MU", "Hu": "HU", "Ct": "CT", "Mr": "MR", "Cbct": "CBCT",
    "Pet": "PET", "Mbs": "MBS (model-based segmentation)", "Uid": "UID", "Db": "DB", "Ptv": "PTV",
    "Ctv": "CTV", "Gtv": "GTV", "Oar": "OAR", "Ui": "UI", "Id": "ID", "Sop": "SOP",
}

CATEGORIES = [
    ("create", "생성/추가", "Create / add", ["Create", "Add", "Make", "Generate", "New", "Append", "Insert"]),
    ("delete", "삭제/제거", "Delete / remove", ["Delete", "Remove", "Clear", "Unload", "Discard"]),
    ("compute", "계산/최적화/평가", "Compute / optimize / evaluate",
     ["Compute", "Calculate", "Recompute", "Run", "Execute", "Optimize", "Reoptimize", "Evaluate",
      "Estimate", "Simulate", "Predict", "Train"]),
    ("query", "조회/검사", "Query / check",
     ["Get", "Query", "Find", "Search", "List", "Is", "Has", "Can", "Check", "Validate", "Verify"]),
    ("edit", "설정/편집", "Set / edit",
     ["Set", "Edit", "Update", "Change", "Modify", "Assign", "Rename", "Reset", "Move", "Rotate",
      "Translate", "Scale", "Rescale", "Normalize", "Toggle", "Select", "Apply", "Enable", "Disable",
      "Replace", "Sort", "Link", "Unlink"]),
    ("io", "가져오기/내보내기/저장", "Import / export / save",
     ["Import", "Export", "Load", "Save", "Store", "Send", "Print", "Open", "Close", "Show", "Hide",
      "Refresh"]),
    ("copy", "복사/변환/정합", "Copy / convert / register",
     ["Copy", "Duplicate", "Convert", "Transform", "Merge", "Split", "Mirror", "Resample", "Deform",
      "Register", "Adapt"]),
    ("approve", "승인/잠금", "Approve / lock",
     ["Approve", "Unapprove", "Lock", "Unlock", "Accept", "Reject", "Cancel"]),
    ("geometry", "형상 처리", "Geometry processing",
     ["Simplify", "Smooth", "Expand", "Contract", "Interpolate", "Crop", "Fill", "Threshold", "Grow",
      "Fit", "Segment", "Round", "Snap"]),
]
_VERB_TO_CAT = {}
_CAT_LABEL = {"other": (S["cat_other"][0], S["cat_other"][1])}
for _cid, _ko, _en, _verbs in CATEGORIES:
    _CAT_LABEL[_cid] = (_ko, _en)
    for _v in _verbs:
        _VERB_TO_CAT[_v] = _cid


def cat_label(cid):
    pair = _CAT_LABEL.get(cid, _CAT_LABEL["other"])
    return pair[0] if LANG == "ko" else pair[1]


def split_camel(name):
    toks = re.findall(r"[A-Z]+(?=[A-Z][a-z]|\d|_|$)|[A-Z]?[a-z]+|\d+", name.replace("_", " "))
    return toks or [name]


def _terms_ko(tokens):
    out, i = [], 0
    while i < len(tokens):
        hit = None
        for span in (3, 2, 1):
            if i + span <= len(tokens):
                key = "".join(t[:1].upper() + t[1:] for t in tokens[i:i + span])
                if key in TERM_KO:
                    hit = (TERM_KO[key], span)
                    break
        if hit:
            out.append(hit[0])
            i += hit[1]
        else:
            out.append(tokens[i])
            i += 1
    return " ".join(out)


def _terms_en(tokens):
    out = []
    for t in tokens:
        key = t[:1].upper() + t[1:]
        if key in TERM_EN:
            out.append(TERM_EN[key])
        elif t.isupper() and len(t) > 1:
            out.append(t)
        else:
            out.append(t.lower())
    return " ".join(out)


def _third_person(v):
    lv = v.lower()
    if lv.endswith(("s", "sh", "ch", "x", "z")):
        return v + "es"
    if lv.endswith("y") and len(lv) > 1 and lv[-2] not in "aeiou":
        return v[:-1] + "ies"
    return v + "s"


def method_verb(name):
    toks = split_camel(name)
    v = toks[0][:1].upper() + toks[0][1:] if toks else ""
    return v if v in VERB_KO else None


def method_category(name):
    v = method_verb(name)
    return _VERB_TO_CAT.get(v, "other") if v else "other"


def annotate_method(name):
    toks = split_camel(name)
    verb = method_verb(name)
    if LANG == "ko":
        if verb:
            obj = _terms_ko(toks[1:]) or tr("r_target_obj")
            return "[%s] %s 을(를) %s" % (VERB_KO[verb], obj, VERB_KO[verb])
        return "%s 관련 기능" % _terms_ko(toks)
    obj = _terms_en(toks[1:]) or tr("r_target_obj")
    if verb in ("Is", "Has", "Can"):
        return "Checks whether it %s %s" % (verb.lower(), obj)
    if verb:
        return "%s %s" % (_third_person(verb), obj)
    return "Function related to %s" % _terms_en(toks)


def annotate_property(name):
    toks = split_camel(name)
    if LANG == "ko":
        if toks and toks[0] in ("Is", "Has", "Can"):
            return "%s 여부 (True/False)" % _terms_ko(toks[1:])
        return _terms_ko(toks)
    if toks and toks[0] in ("Is", "Has", "Can"):
        return "Whether %s %s (True/False)" % (toks[0].lower(), _terms_en(toks[1:]))
    t = _terms_en(toks)
    return t[:1].upper() + t[1:]


# ---------------------------------------------------------------------------
# docstring / signature / parameter parsing
# ---------------------------------------------------------------------------
def get_doc(name, fn):
    doc = ""
    try:
        doc = fn.__doc__ or ""
    except Exception:
        doc = ""
    doc = to_text(doc) if doc else ""
    generic = ""
    try:
        generic = to_text(type(fn).__doc__ or "")
    except Exception:
        pass
    if generic and doc.strip() == generic.strip():
        doc = ""
    if not doc.strip():
        try:
            import pydoc
            try:
                d = pydoc.render_doc(fn, renderer=pydoc.plaintext)
            except TypeError:
                d = pydoc.render_doc(fn)
            d = re.sub(".\b", "", to_text(d))
            head = "\n".join(d.splitlines()[:4])
            if name in head and (not generic or generic.strip() not in d):
                doc = d
        except Exception:
            pass
    try:
        doc = inspect.cleandoc(doc)
    except Exception:
        pass
    return doc.strip()


def get_signature(name, fn, doc):
    try:
        if hasattr(inspect, "signature"):
            s = to_text(inspect.signature(fn))
            if s not in ("(*args, **kwargs)", "(*args, **kw)", "(*args)", "(**kwargs)"):
                return name + s
    except Exception:
        pass
    m = re.search(r"^\s*(?:[\w\.]+\.)?" + re.escape(name) + r"\s*\(([^)]*)\)", doc or "", re.M)
    if m:
        return "%s(%s)" % (name, " ".join(m.group(1).split()))
    return None


def _split_args(inner):
    parts, depth, cur = [], 0, ""
    for ch in inner:
        if ch in "([{<":
            depth += 1
        elif ch in ")]}>":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        parts.append(cur)
    return [p.strip() for p in parts if p.strip()]


_P_SIG = re.compile(r"^\**\s*(\w+)\s*(?::\s*([^=]+?))?\s*(?:=\s*(.+))?$")
_H_PARAMS = re.compile(r"^\s*(parameters|parameter|arguments|args|params|inputs?|"
                       r"keyword arguments|매개변수|인자)\s*:?\s*$", re.I)
_H_RETURNS = re.compile(r"^\s*(returns?|return value|반환값?)\s*:?\s*(.*)$", re.I)
_H_OTHER = re.compile(r"^\s*(raises|exceptions?|examples?|notes?|see also|remarks?|"
                      r"example usage|usage)\s*:?\s*$", re.I)
_UNDER = re.compile(r"^\s*[-=~]{3,}\s*$")
_P1 = re.compile(r"^\s*[-*•]?\s*(\w+)\s*\(([^)]*)\)\s*[:\-–]\s*(.*)$")
_P2 = re.compile(r"^\s*[-*•]?\s*(\w+)\s+:\s*(\S+(?:\s*,\s*optional)?)\s*$")
_P3 = re.compile(r"^\s*[-*•]?\s*(\w+)\s*[:\-–]\s+(.+)$")
_SPH_P = re.compile(r"^\s*:param\s+(?:([\w\.\[\], ]+?)\s+)?(\w+)\s*:\s*(.*)$")
_SPH_T = re.compile(r"^\s*:type\s+(\w+)\s*:\s*(.*)$")
_SPH_R = re.compile(r"^\s*:returns?:\s*(.*)$")
_SPH_RT = re.compile(r"^\s*:rtype:\s*(.*)$")


def parse_params(doc, sig):
    order, params = [], {}

    def get(name):
        if name not in params:
            params[name] = {"name": name, "type": "", "default": None, "desc": ""}
            order.append(name)
        return params[name]

    if sig and "(" in sig:
        inner = sig[sig.find("(") + 1:sig.rfind(")")]
        for tok in _split_args(inner):
            if tok in ("self", "/", "*", "cls"):
                continue
            m = _P_SIG.match(tok)
            if not m:
                continue
            p = get(m.group(1))
            if m.group(2):
                p["type"] = m.group(2).strip()
            if m.group(3) is not None:
                p["default"] = m.group(3).strip()

    returns = []
    section, last, base = None, None, None
    for line in (doc or "").splitlines():
        m = _SPH_P.match(line)
        if m:
            p = get(m.group(2))
            if m.group(1):
                p["type"] = m.group(1).strip()
            p["desc"] = m.group(3).strip()
            last, section = p, "sphinx"
            continue
        m = _SPH_T.match(line)
        if m:
            get(m.group(1))["type"] = m.group(2).strip()
            continue
        m = _SPH_R.match(line) or _SPH_RT.match(line)
        if m:
            returns.append(m.group(1).strip())
            continue
        if _H_PARAMS.match(line):
            section, last, base = "params", None, None
            continue
        m = _H_RETURNS.match(line)
        if m and (not m.group(2) or section != "params" or len(line) - len(line.lstrip()) == 0):
            section = "returns"
            if m.group(2).strip():
                returns.append(m.group(2).strip())
            continue
        if _H_OTHER.match(line):
            section = None
            continue
        if _UNDER.match(line):
            continue
        if section == "params":
            if not line.strip():
                continue
            indent = len(line) - len(line.lstrip())
            if base is not None and indent > base and last is not None:
                last["desc"] = (last["desc"] + " " + line.strip()).strip()
                continue
            m1, m2, m3 = _P1.match(line), _P2.match(line), _P3.match(line)
            if m1:
                p = get(m1.group(1))
                p["type"] = p["type"] or m1.group(2).strip()
                p["desc"] = m1.group(3).strip()
            elif m2:
                p = get(m2.group(1))
                p["type"] = p["type"] or m2.group(2).strip()
            elif m3:
                p = get(m3.group(1))
                p["desc"] = m3.group(2).strip()
            elif last is not None:
                last["desc"] = (last["desc"] + " " + line.strip()).strip()
                continue
            else:
                continue
            last = p
            if base is None:
                base = indent
        elif section == "sphinx":
            if line.strip() and last is not None and line.startswith((" ", "\t")):
                last["desc"] = (last["desc"] + " " + line.strip()).strip()
        elif section == "returns":
            if line.strip():
                returns.append(line.strip())
            elif returns:
                section = None

    result = []
    for n in order:
        p = params[n]
        low = (p["type"] + " " + p["desc"]).lower()
        p["optional"] = bool(p["default"] is not None or "optional" in low or "선택" in low)
        result.append(p)
    return result, " ".join(returns).strip()


def first_paragraph(doc, sig_name=None, limit=500):
    if not doc:
        return ""
    for p in re.split(r"\n\s*\n", doc.strip()):
        t = " ".join(p.split())
        if sig_name and re.match(r"^(?:[\w\.]+\.)?" + re.escape(sig_name) + r"\s*\(", t):
            rest = re.sub(r"^(?:[\w\.]+\.)?" + re.escape(sig_name) + r"\s*\([^)]*\)\s*", "", t)
            if not rest:
                continue
            t = rest
        if _H_PARAMS.match(t) or _H_RETURNS.match(t):
            break
        return t if len(t) <= limit else t[:limit - 1] + "…"
    return ""


def describe_method(name, fn):
    doc = get_doc(name, fn)
    sig = get_signature(name, fn, doc)
    params, returns = parse_params(doc, sig)
    return {
        "name": name, "signature": sig, "doc": doc, "summary": first_paragraph(doc, name),
        "params": params, "returns": returns,
        "deprecated": bool(DEPRECATED_RE.search(doc or "")),
        "category": method_category(name), "annotation": annotate_method(name),
    }


def failed_method(name, e):
    return {"name": name, "signature": None, "doc": "", "summary": "", "params": [],
            "returns": "", "deprecated": False, "category": method_category(name),
            "annotation": annotate_method(name), "error": err_text(e)}


def is_method(val):
    if val is None or isinstance(val, string_types + NUMBER_TYPES + (bool, dict, list, tuple)):
        return False
    try:
        if not callable(val):
            return False
    except Exception:
        return False
    return not has_api_members(val)


# ---------------------------------------------------------------------------
# RayStation version / build type detection
# ---------------------------------------------------------------------------
VERSION_RE = re.compile(
    r"(?<![0-9A-Za-z])((?:20[1-4]\d|1\d)[AB])"
    r"(?:[\s_\-]*(SP\s?\d+))?(?:[\s_\-]*(R|DR)(?![A-Za-z0-9]))?", re.I)
VERSION_RE2 = re.compile(r"RayStation[\s_\-]*(20[1-4]\d)(?![0-9])", re.I)
PRODVER_RE = re.compile(r"(?<![\d.])(\d+\.\d+\.\d+(?:\.\d+)?)(?![\d.])")

BUILD_RULES = [   # priority order
    ("evaluation", re.compile(r"\bevaluation\b|\beval\b|\btrial\b|\bdemo\b", re.I)),
    ("research", re.compile(r"\bresearch\b|(?<![0-9A-Za-z])(?:20[1-4]\d|1\d)[AB]"
                            r"(?:[\s_\-]*SP\s?\d+)?[\s_\-]*R(?![A-Za-z0-9])", re.I)),
    ("nonclinical", re.compile(r"not\s+for\s+clinical|non[\s\-]?clinical", re.I)),
    ("validation", re.compile(r"\bvalidation\b|\bbeta\b|\bpreview\b|pre[\s\-]?release|"
                              r"release\s+candidate", re.I)),
    ("training", re.compile(r"\btraining\b|\beducation(?:al)?\b", re.I)),
    ("clinical", re.compile(r"\bclinical\b", re.I)),
]


def _title_keywords(title):
    """Keep only version/build keywords from a window title (no patient data)."""
    found = []
    for m in VERSION_RE.finditer(title):
        found.append(" ".join(x for x in m.groups() if x))
    if not found:
        for m in VERSION_RE2.finditer(title):
            found.append(m.group(1))
    for bt, rx in BUILD_RULES:
        m = rx.search(title)
        if m:
            found.append(m.group(0))
    if re.search(r"raystation", title, re.I):
        found.insert(0, "RayStation")
    out = []
    for f in found:
        if f not in out:
            out.append(f)
    return " · ".join(out)


def collect_env_evidence():
    ev = []       # (source, text, used_for_detection)
    procs = []
    if connect is not None:
        p = getattr(connect, "__file__", None)
        if p:
            ev.append(("connect.__file__", to_text(p), True))
        for name in api_names(connect):
            low = name.lower()
            if "version" in low or "build" in low or "license" in low or "edition" in low:
                try:
                    val = getattr(connect, name)
                    if callable(val):
                        # only argument-less getters such as get_*version* are called
                        if not re.match(r"^get_\w*(version|build|edition)\w*$", name, re.I):
                            continue
                        val = val()
                    ev.append(("connect.%s" % name, short(val, 200), True))
                except Exception:
                    pass
    ev.append(("sys.executable", to_text(sys.executable), True))
    for k in sorted(os.environ.keys()):
        if re.search(r"RAYSTATION|RAYSEARCH|^RSL", k, re.I):
            ev.append(("env:%s" % k, short(os.environ[k], 200), True))
    try:
        import clr  # noqa: F401
        from System.Diagnostics import Process
        from System.IO import File
        for pr in Process.GetProcesses():
            try:
                pname = to_text(pr.ProcessName)
            except Exception:
                continue
            if "raystation" not in pname.lower():
                continue
            info = {"name": pname}
            try:
                t = to_text(pr.MainWindowTitle)
                if t:
                    info["title_keywords"] = _title_keywords(t)
            except Exception:
                pass
            try:
                mm = pr.MainModule
                fn = to_text(mm.FileName)
                info["path"] = fn
                fvi = mm.FileVersionInfo
                for attr in ("ProductName", "ProductVersion", "FileVersion", "FileDescription",
                             "Comments", "SpecialBuild", "PrivateBuild"):
                    try:
                        v = getattr(fvi, attr)
                        if v:
                            info[attr] = to_text(v)
                    except Exception:
                        pass
                for attr in ("IsPreRelease", "IsDebug", "IsSpecialBuild", "IsPrivateBuild"):
                    try:
                        info[attr] = bool(getattr(fvi, attr))
                    except Exception:
                        pass
                try:
                    info["mtime"] = to_text(File.GetLastWriteTime(fn).ToString("yyyy-MM-dd HH:mm:ss"))
                except Exception:
                    pass
            except Exception as e:
                info["module_error"] = err_text(e)
            procs.append(info)
    except Exception:
        pass
    return ev, procs


def _path_affinity(proc_path, refs):
    """How much a RayStation process path shares its install folder with our process."""
    if not proc_path:
        return 0
    d = os.path.dirname(proc_path).lower()
    best = 0
    for r in refs:
        r = (r or "").lower()
        n = 0
        for a, b in zip(d, r):
            if a != b:
                break
            n += 1
        best = max(best, n)
    return best


def detect_environment():
    ev, procs = collect_env_evidence()
    refs = [t for s, t, _u in ev if s in ("connect.__file__", "sys.executable")]
    procs.sort(key=lambda p: -_path_affinity(p.get("path"), refs))
    main = procs[0] if procs else None

    texts = []   # (source, text) used for detection
    if main:
        for k in ("title_keywords", "ProductName", "FileDescription", "Comments", "SpecialBuild",
                  "PrivateBuild", "path", "ProductVersion", "FileVersion"):
            if main.get(k):
                texts.append(("RayStation process: %s" % k, main[k]))
    for s, t, used in ev:
        if used:
            texts.append((s, t))

    version, version_src = None, None
    for s, t in texts:
        m = VERSION_RE.search(t)
        if m:
            version = m.group(1).upper()
            if m.group(2):
                version += " " + m.group(2).upper().replace(" ", "")
            if m.group(3):
                version += "-" + m.group(3).upper()
            version_src = s
            break
        m = VERSION_RE2.search(t)
        if m:
            version, version_src = m.group(1), s
            break

    prodver = None
    if main:
        for k in ("ProductVersion", "FileVersion"):
            m = PRODVER_RE.search(main.get(k, "") or "")
            if m:
                prodver = m.group(1)
                break

    flags, basis = [], []
    for s, t in texts:
        for bt, rx in BUILD_RULES:
            m = rx.search(t)
            if not m:
                continue
            if bt == "clinical" and BUILD_RULES[2][1].search(t):
                continue
            if bt not in flags:
                flags.append(bt)
            basis.append("%s → \"%s\"" % (s, short(m.group(0), 40)))
    if main and main.get("IsPreRelease"):
        if "validation" not in flags:
            flags.append("validation")
        basis.append("FileVersionInfo.IsPreRelease = True")

    build = None
    for bt, _rx in BUILD_RULES:
        if bt in flags:
            build = bt
            break
    if build is None:
        build = "clinical_assumed" if (version or prodver or main) else "unknown"

    target_ok = None
    if version:
        a = re.sub(r"[^0-9A-Z]", "", version.upper())
        b = re.sub(r"[^0-9A-Z]", "", SCRIPT_TARGET_VERSION.upper())
        target_ok = a.startswith(b) or b.startswith(a)

    label = version or SCRIPT_TARGET_VERSION
    return {
        "version": version, "version_source": version_src, "product_version": prodver,
        "build_type": build, "build_flags": flags, "build_basis": basis,
        "target_version": SCRIPT_TARGET_VERSION, "target_match": target_ok,
        "exe_path": main.get("path") if main else None,
        "exe_mtime": main.get("mtime") if main else None,
        "processes": procs, "evidence": [(s, t) for s, t, _u in ev],
        "default_label": label,
    }


def build_type_text(bt):
    return tr("bt_" + bt) if ("bt_" + bt) in S else bt


def build_short(bt):
    return {"clinical": "Clinical", "clinical_assumed": "Clinical", "research": "Research",
            "evaluation": "Evaluation", "nonclinical": "NonClinical", "validation": "Validation",
            "training": "Training"}.get(bt, "")


def env_summary(env):
    parts = ["RayStation " + (env.get("version") or tr("r_nover"))]
    if env.get("product_version"):
        parts.append(env["product_version"])
    parts.append(build_type_text(env["build_type"]))
    return " · ".join(parts)


# ---------------------------------------------------------------------------
# API walker
# ---------------------------------------------------------------------------
class TypeInfo(object):
    def __init__(self, key, expr, depth, type_name):
        self.key = key
        self.type_name = type_name
        self.first_expr = expr
        self.paths = [normalize_path(expr)]
        self.depth = depth
        self.properties = []
        self.methods = []
        self.status = "pending"
        self.display = None

    def to_dict(self):
        return {"key": self.key, "display": self.display, "type_name": self.type_name,
                "paths": self.paths, "example_expr": root_code(self.first_expr),
                "status": self.status, "properties": self.properties, "methods": self.methods}


class ApiWalker(object):
    def __init__(self, opts, log=None, progress=None):
        self.opts = opts
        self.log = log or (lambda m: None)
        self.progress = progress or (lambda i, n, t: None)
        self.types = {}
        self.order = []
        self.queue = []
        self.cancel = False

    def register(self, obj, expr, depth):
        key = sig_key(api_names(obj))
        ti = self.types.get(key)
        norm = normalize_path(expr)
        if ti is not None:
            if norm not in ti.paths and len(ti.paths) < 8:
                ti.paths.append(norm)
            return key
        ti = TypeInfo(key, expr, depth, display_type_name(obj))
        self.types[key] = ti
        self.order.append(key)
        if depth > self.opts["max_depth"]:
            ti.status = "depth_limit"
        elif len(self.types) > self.opts["max_types"]:
            ti.status = "type_limit"
        else:
            self.queue.append((key, obj))
        return key

    def as_sequence(self, val):
        if isinstance(val, string_types) or isinstance(val, dict):
            return None
        k = self.opts["samples"]
        if isinstance(val, (list, tuple)):
            return len(val), list(val[:k])
        try:
            n = len(val)
        except Exception:
            return None
        items = []
        try:
            for i in range(min(n, k)):
                items.append(val[i])
        except Exception:
            items = []
            try:
                it = iter(val)
                for i in range(min(n, k)):
                    items.append(next(it))
            except Exception:
                return None
        return n, items

    def classify(self, val, expr, depth, nest=0):
        mask = self.opts.get("mask", True)
        if val is None:
            return ("none", "None", "None", [])
        if isinstance(val, bool):
            return ("primitive", "bool", to_text(val), [])
        if isinstance(val, NUMBER_TYPES):
            return ("primitive", type(val).__name__, short(val), [])
        if isinstance(val, string_types):
            t = to_text(val)
            return ("primitive", "str", tr("p_masked", len(t)) if mask else '"%s"' % short(t, 50), [])
        if isinstance(val, dict):
            keys = [to_text(k) for k in list(val.keys())[:12]]
            more = ", …" if len(val) > 12 else ""
            has_text = any(isinstance(v, string_types) for v in list(val.values()))
            sample = tr("p_masked2") if (mask and has_text) else short(val)
            return ("dict", "dict{%s%s}" % (", ".join(keys), more), sample, [])
        tn = dotnet_type_name(val)
        if tn and tn.startswith("System.") and not tn.endswith("]") and "Collections" not in tn:
            return ("primitive", tn, tr("p_masked2") if mask else short(val), [])
        seq = self.as_sequence(val)
        if seq is not None:
            n, items = seq
            if n == 0:
                return ("empty", tr("p_empty_t"), tr("p_count", 0), [])
            if nest > 3:
                return ("collection", "Collection", tr("p_count", n), [])
            descs, targets = [], []
            for idx, item in enumerate(items):
                _k, td, _s, tg = self.classify(item, "%s[%d]" % (expr, idx), depth, nest + 1)
                if td not in descs:
                    descs.append(td)
                for t in tg:
                    if t not in targets:
                        targets.append(t)
            return ("collection", "Collection<%s>" % " | ".join(descs), tr("p_count", n), targets)
        if has_api_members(val):
            key = self.register(val, expr, depth)
            return ("object", "@@%s" % key, "", [key])
        try:
            if callable(val):
                return ("callable", type(val).__name__, "", [])
        except Exception:
            pass
        return ("primitive", tn or type(val).__name__, short(val), [])

    def explore(self, ti, obj):
        for name in sorted(api_names(obj)):
            if self.cancel:
                break
            if name in self.opts["skip"]:
                ti.properties.append({"name": name, "kind": "skipped", "type": tr("p_skipped_t"),
                                      "sample": "", "targets": [], "note": tr("p_skipped_n"),
                                      "annotation": annotate_property(name)})
                continue
            try:
                val = getattr(obj, name)
            except Exception as e:
                ti.properties.append({"name": name, "kind": "error", "type": tr("p_error_t"),
                                      "sample": "", "targets": [], "note": err_text(e),
                                      "annotation": annotate_property(name)})
                continue
            if is_method(val):
                try:
                    ti.methods.append(describe_method(name, val))
                except Exception as e:
                    ti.methods.append(failed_method(name, e))
                continue
            try:
                kind, tdesc, sample, targets = self.classify(
                    val, "%s.%s" % (ti.first_expr, name), ti.depth + 1)
                note = ""
            except Exception as e:
                kind, tdesc, sample, targets, note = "error", tr("p_fail_t"), "", [], err_text(e)
            ti.properties.append({"name": name, "kind": kind, "type": tdesc, "sample": sample,
                                  "targets": targets, "note": note,
                                  "annotation": annotate_property(name)})
        ti.status = "explored" if not self.cancel else "partial"

    def run(self, roots):
        for label, obj in roots:
            self.register(obj, label, 0)
        i = 0
        while i < len(self.queue):
            if self.cancel:
                break
            key, obj = self.queue[i]
            i += 1
            ti = self.types[key]
            self.progress(i, len(self.queue), ti)
            try:
                self.explore(ti, obj)
            except Exception as e:
                ti.status = "error"
                self.log(tr("log_fail_type", ti.first_expr, err_text(e)))
        for k in self.order:
            if self.types[k].status == "pending":
                self.types[k].status = "cancelled"
        self.assign_display_names()

    def assign_display_names(self):
        count = {}
        for k in self.order:
            n = self.types[k].type_name
            if n:
                count[n] = count.get(n, 0) + 1
        used = {}
        for k in self.order:
            ti = self.types[k]
            name = ti.type_name if (ti.type_name and count.get(ti.type_name) == 1) else None
            if not name:
                seg = ti.paths[0].split(".")[-1]
                name = tr("elem", seg[:-2]) if seg.endswith("[]") else seg
                if ti.type_name:
                    name = "%s (%s)" % (name, ti.type_name)
            base, j = name, 2
            while name in used:
                name = "%s #%d" % (base, j)
                j += 1
            used[name] = True
            ti.display = name


# ---------------------------------------------------------------------------
# Roots / global functions
# ---------------------------------------------------------------------------
def acquire_roots(root_names, log):
    objs, status = [], []
    if connect is None:
        for r in root_names:
            status.append((r, "fail", tr("no_connect")))
        return objs, status
    for r in root_names:
        try:
            o = connect.get_current(r)
            if o is None:
                raise Exception(tr("none_ret"))
            objs.append((r, o))
            status.append((r, "ok", ""))
            log("  ✓ get_current(\"%s\")" % r)
        except Exception as e:
            status.append((r, "fail", err_text(e)))
            log("  ✗ get_current(\"%s\") - %s" % (r, err_text(e)))
    return objs, status


def collect_globals():
    funcs, others = [], []
    if connect is None:
        return funcs, others
    import types as _types
    for name in sorted(api_names(connect)):
        try:
            val = getattr(connect, name)
        except Exception as e:
            others.append({"name": name, "type": "?", "note": err_text(e)})
            continue
        if isinstance(val, _types.ModuleType):
            continue
        try:
            is_call = callable(val)
        except Exception:
            is_call = False
        if is_call:
            try:
                funcs.append(describe_method(name, val))
            except Exception as e:
                funcs.append(failed_method(name, e))
        else:
            others.append({"name": name, "type": type(val).__name__, "note": short(val)})
    return funcs, others


# ---------------------------------------------------------------------------
# Markdown writer
# ---------------------------------------------------------------------------
KIND_ICON = {"error": "🚫", "skipped": "⏭", "empty": "∅"}


def status_text(st):
    k = "st_" + st
    return tr(k) if k in S else st


class MarkdownWriter(object):
    def __init__(self, walker, meta, env, globals_, root_status):
        self.w = walker
        self.meta = meta
        self.env = env
        self.funcs, self.gothers = globals_
        self.root_status = root_status
        self.lines = []

    def link(self, s):
        def rep(m):
            ti = self.w.types.get(m.group(1))
            return "[%s](#t-%s)" % (ti.display if ti else m.group(1), m.group(1))
        return re.sub(r"@@(t[0-9a-f]{10})", rep, s)

    def a(self, s=""):
        self.lines.append(to_text(s))

    @staticmethod
    def manchor(tkey, mname):
        return "m-%s-%s" % (tkey, mname)

    def render_header(self):
        a, meta, env = self.a, self.meta, self.env
        a("# " + tr("r_title", meta["label"]))
        a()
        a("> " + tr("r_intro"))
        a()
        a("| %s | %s |" % (tr("r_item"), tr("r_value")))
        a("|---|---|")
        a("| **%s** | **%s** |" % (tr("r_created"), meta["generated"]))
        a("| %s | **%s** |" % (tr("r_label"), md_escape(meta["label"])))
        a("| %s | %s |" % (tr("r_target"), SCRIPT_TARGET_VERSION))
        if env.get("version"):
            check = tr("r_match") if env.get("target_match") else tr("r_mismatch", SCRIPT_TARGET_VERSION)
            a("| %s | **%s** — %s |" % (tr("r_detver"), md_escape(env["version"]), check))
        else:
            a("| %s | %s |" % (tr("r_detver"), tr("r_nover")))
        if env.get("product_version"):
            a("| %s | %s |" % (tr("r_prodver"), md_escape(env["product_version"])))
        a("| **%s** | **%s** |" % (tr("r_buildtype"), md_escape(build_type_text(env["build_type"]))))
        if env.get("build_basis"):
            a("| %s | %s |" % (tr("r_basis"), md_escape("\n".join(env["build_basis"][:8]))))
        if env.get("exe_path"):
            a("| %s | `%s` |" % (tr("r_exe"), md_escape(env["exe_path"])))
        if env.get("exe_mtime"):
            a("| %s | %s |" % (tr("r_exe_mtime"), env["exe_mtime"]))
        a("| %s | %s |" % (tr("r_python"), md_escape(meta["python"])))
        a("| %s | %s |" % (tr("r_script"), md_escape(SCRIPT_TITLE)))
        a("| %s | %s |" % (tr("r_settings"), tr("r_settings_v", meta["max_depth"], meta["max_types"],
                                                 meta["samples"], tr("r_on") if meta["mask"] else tr("r_off"))))
        a("| %s | %s |" % (tr("r_elapsed"), tr("r_sec", meta["elapsed"])))
        a()

    def render_env(self):
        a, env = self.a, self.env
        a("<details><summary>%s</summary>" % tr("r_env_h"))
        a()
        a(tr("r_env_note"))
        a()
        a("| %s | %s |" % (tr("r_source"), tr("r_value")))
        a("|---|---|")
        for p in env.get("processes", []):
            for k in sorted(p.keys()):
                a("| RayStation process `%s`: %s | %s |" % (md_escape(p.get("name")), k, md_escape(p[k])))
        for s, t in env.get("evidence", []):
            a("| %s | %s |" % (md_escape(s), md_escape(t)))
        a()
        a("</details>")
        a()

    def render(self):
        w = self.w
        types = [w.types[k] for k in w.order]
        n_m = sum(len(t.methods) for t in types)
        n_p = sum(len(t.properties) for t in types)
        dep = [(t, m) for t in types for m in t.methods if m.get("deprecated")]
        errs = [(t, p) for t in types for p in t.properties if p["kind"] == "error"]
        empties = [(t, p) for t in types for p in t.properties if p["kind"] == "empty"]
        unexplored = [t for t in types if t.status not in ("explored", "partial")]
        a = self.a

        self.render_header()
        self.render_env()

        a("### " + tr("r_roots_h"))
        a()
        a("| %s | %s | %s |" % (tr("r_root"), tr("r_result"), tr("r_note")))
        a("|---|---|---|")
        for r, st, msg in self.root_status:
            a("| `get_current(\"%s\")` | %s | %s |" % (r, tr("r_ok") if st == "ok" else tr("r_fail"),
                                                      md_escape(msg)))
        a()
        a("## " + tr("r_summary"))
        a()
        a("| %s | %s |" % (tr("r_kind"), tr("r_count")))
        a("|---|---|")
        for key, n in (("r_n_types", len(types)), ("r_n_methods", n_m), ("r_n_globals", len(self.funcs)),
                       ("r_n_props", n_p), ("r_n_dep", len(dep)), ("r_n_err", len(errs)),
                       ("r_n_empty", len(empties)), ("r_n_unexp", len(unexplored))):
            a("| %s | %d |" % (tr(key), n))
        a()
        a(tr("r_legend"))
        a()

        a("## " + tr("r_toc"))
        a()
        a("- [%s](#sec-index)" % tr("r_idx_h"))
        a("- [%s](#sec-warn)" % tr("r_warn_h"))
        a("- [%s](#sec-globals)" % tr("r_glob_h"))
        a("- [%s](#sec-types)" % tr("r_types_h"))
        for t in types:
            a("  - [%s](#t-%s) — %s%s" % (t.display, t.key, tr("r_n_mp", len(t.methods), len(t.properties)),
                                        "" if t.status == "explored" else " · " + status_text(t.status)))
        a()

        # feature index
        a('<a id="sec-index"></a>')
        a()
        a("## " + tr("r_idx_h"))
        a()
        a(tr("r_idx_note"))
        a()
        buckets = {}
        for t in types:
            for m in t.methods:
                buckets.setdefault(m["category"], []).append((t, m))
        cat_order = [c[0] for c in CATEGORIES] + ["other"]
        a("| %s | %s |" % (tr("r_cat"), tr("r_n_m")))
        a("|---|---|")
        for c in cat_order:
            if c in buckets:
                a("| %s | %d |" % (cat_label(c), len(buckets[c])))
        a()
        for c in cat_order:
            if c not in buckets:
                continue
            a("### %s (%d)" % (cat_label(c), len(buckets[c])))
            a()
            for t, m in sorted(buckets[c], key=lambda x: (x[1]["name"], x[0].display)):
                a("- [`%s`](#%s) · %s — %s%s" % (m["name"], self.manchor(t.key, m["name"]), t.display,
                                               m["annotation"], " ⚠️" if m.get("deprecated") else ""))
            a()

        # watch list
        a('<a id="sec-warn"></a>')
        a()
        a("## " + tr("r_warn_h"))
        a()
        a("### " + tr("r_dep_h"))
        a()
        if dep:
            for t, m in dep:
                a("- [`%s.%s`](#%s)" % (t.display, m["name"], self.manchor(t.key, m["name"])))
        else:
            a("- " + tr("r_dep_none"))
        a()
        a("### " + tr("r_err_h"))
        a()
        a(tr("r_err_note"))
        a()
        if errs:
            a("| %s | %s | %s |" % (tr("r_where"), tr("r_prop"), tr("r_error")))
            a("|---|---|---|")
            for t, p in errs:
                a("| [%s](#t-%s) | `%s` | %s |" % (t.display, t.key, p["name"], md_escape(p["note"])))
        else:
            a("- " + tr("r_none"))
        a()
        a("### " + tr("r_empty_h"))
        a()
        a(tr("r_empty_note"))
        a()
        if empties:
            for t, p in empties:
                a("- `%s.%s`" % (t.paths[0], p["name"]))
        else:
            a("- " + tr("r_none"))
        a()
        if unexplored:
            a("### " + tr("r_unexp_h"))
            a()
            for t in unexplored:
                a("- [%s](#t-%s) — %s (`%s`)" % (t.display, t.key, status_text(t.status), t.paths[0]))
            a()

        # globals
        a('<a id="sec-globals"></a>')
        a()
        a("## " + tr("r_glob_h"))
        a()
        a(tr("r_glob_note"))
        a()
        if not self.funcs and not self.gothers:
            a("- " + tr("r_glob_fail"))
        for m in self.funcs:
            self.render_method("connect", m, "connect")
        if self.gothers:
            a("#### " + tr("r_glob_other"))
            a()
            a("| %s | %s | %s |" % (tr("r_name"), tr("r_type"), tr("r_valnote")))
            a("|---|---|---|")
            for o in self.gothers:
                a("| `%s` | %s | %s |" % (o["name"], md_escape(o["type"]), md_escape(o["note"])))
            a()

        # types
        a('<a id="sec-types"></a>')
        a()
        a("## " + tr("r_types_h"))
        a()
        for t in types:
            self.render_type(t)
        return "\n".join(self.lines) + "\n"

    def render_type(self, t):
        a = self.a
        a('<a id="t-%s"></a>' % t.key)
        a()
        a("### %s" % t.display)
        a()
        a("- **%s**: %s" % (tr("r_paths"), ", ".join("`%s`" % p for p in t.paths)))
        a("- **%s**: `%s`" % (tr("r_code"), root_code(t.first_expr)))
        if t.type_name:
            a("- **%s**: `%s`" % (tr("r_internal"), t.type_name))
        a("- **%s**: %s · %s" % (tr("r_state"), status_text(t.status),
                                 tr("r_n_mp", len(t.methods), len(t.properties))))
        a()
        if t.properties:
            a("#### " + tr("r_props_h"))
            a()
            a(tr("r_props_cols"))
            a("|---|---|---|---|")
            for p in t.properties:
                icon = KIND_ICON.get(p["kind"], "")
                sample = p["note"] if p["kind"] in ("error", "skipped") else p["sample"]
                a("| %s`%s` | %s | %s | %s |" % ((icon + " ") if icon else "", p["name"],
                                                 self.link(md_escape(p["type"])), md_escape(sample),
                                                 md_escape(p["annotation"])))
            a()
        if t.methods:
            a("#### " + tr("r_methods_h"))
            a()
            a(tr("r_methods_cols"))
            a("|---|---|---|")
            for m in t.methods:
                a("| [`%s`](#%s)%s | %s | %s |" % (m["name"], self.manchor(t.key, m["name"]),
                                                   " ⚠️" if m.get("deprecated") else "",
                                                   cat_label(m["category"]), md_escape(m["annotation"])))
            a()
            for m in t.methods:
                self.render_method(t.key, m, root_code(t.first_expr))
        a("---")
        a()

    def render_method(self, tkey, m, expr):
        a = self.a
        a('<a id="%s"></a>' % self.manchor(tkey, m["name"]))
        a()
        a("#### `%s`%s" % (m["name"], " ⚠️ Deprecated" if m.get("deprecated") else ""))
        a()
        a("- **%s**: `%s`" % (tr("r_sig"), m.get("signature") or (m["name"] + tr("r_nosig"))))
        a("- **%s**: %s" % (tr("r_auto"), m["annotation"]))
        a("- **%s**: %s" % (tr("r_official"), m.get("summary") or tr("r_nodoc")))
        if m.get("returns"):
            a("- **%s**: %s" % (tr("r_returns"), " ".join(m["returns"].split())))
        if m.get("error"):
            a("- **%s**: %s" % (tr("r_anerr"), m["error"]))
        a()
        if m["params"]:
            a(tr("r_param_cols"))
            a("|---|---|---|---|---|")
            for p in m["params"]:
                a("| `%s` | %s | %s | %s | %s |" % (
                    p["name"], md_escape(p["type"] or "-"),
                    md_escape(p["default"]) if p["default"] is not None else "-",
                    tr("r_opt") if p.get("optional") else tr("r_req"), md_escape(p["desc"] or "-")))
            a()
        call = ("%s(" % m["name"]) if tkey == "connect" else ("%s.%s(" % (expr, m["name"]))
        if m["params"]:
            call += ", ".join("%s=%s" % (p["name"], p["default"] if p["default"] is not None
                                         else "<%s>" % (p["type"] or tr("r_val")))
                              for p in m["params"])
        call += ")"
        a("```python")
        if tkey == "connect":
            a("from connect import *")
        a(call)
        a("```")
        a()
        if m.get("doc"):
            a("<details><summary>%s</summary>" % tr("r_fulldoc"))
            a()
            a("```text")
            a(m["doc"].replace("```", "ˋˋˋ"))
            a("```")
            a()
            a("</details>")
            a()


# ---------------------------------------------------------------------------
# Output path / run
# ---------------------------------------------------------------------------
def safe_filename(s):
    return re.sub(r"[^\w\-\.]+", "_", to_text(s)).strip("_") or "RayStation"


def default_folder():
    home = os.path.expanduser("~")
    for c in (os.path.join(home, "Desktop"), os.path.join(home, "바탕 화면"), home):
        if os.path.isdir(c):
            return c
    return os.getcwd()


def default_file_name(label, env):
    parts = ["RayStation", safe_filename(label)]
    bs = build_short(env.get("build_type", "")) if env else ""
    if bs and bs.lower() not in label.lower():
        parts.append(bs)
    parts.append("API")
    parts.append("KR" if LANG == "ko" else "EN")
    parts.append(datetime.datetime.now().strftime("%Y%m%d_%H%M"))
    return "_".join(parts) + ".md"


def resolve_output_path(text, label, env):
    """Accepts a full file path, a folder, or a name without extension."""
    p = os.path.expandvars(os.path.expanduser(to_text(text).strip().strip('"')))
    if not p:
        raise ValueError(tr("gui_need_path"))
    if os.path.isdir(p) or p.endswith(("\\", "/")):
        p = os.path.join(p, default_file_name(label, env))
    root, ext = os.path.splitext(p)
    if ext.lower() != ".md":
        p = p + ".md" if not ext else root + ".md"
    p = os.path.abspath(p)
    folder = os.path.dirname(p)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)
    return p


def write_text(path, text):
    f = codecs.open(path, "w", "utf-8")
    try:
        f.write(to_text(text))
    finally:
        f.close()


def run_dump(opts, log=None, progress=None, walker_holder=None):
    log = log or (lambda m: print(m))
    t0 = time.time()
    env = opts.get("env") or detect_environment()
    log(tr("log_start", opts["label"]))
    log(tr("log_env", env_summary(env)))
    log(tr("log_roots"))
    roots, root_status = acquire_roots(opts["roots"], log)
    log(tr("log_globals"))
    globals_ = collect_globals()
    log(tr("log_globals_n", len(globals_[0])))

    walker = ApiWalker(opts, log=log, progress=progress)
    if walker_holder is not None:
        walker_holder["walker"] = walker
    log(tr("log_walk"))
    walker.run(roots)
    elapsed = time.time() - t0

    created_dt, created_txt = now_with_offset()
    meta = {
        "label": opts["label"], "language": LANG, "generated": created_txt,
        "generated_iso": created_dt.strftime("%Y-%m-%dT%H:%M:%S"),
        "python": "%s (%s)" % (sys.version.split()[0], sys.platform),
        "script": SCRIPT_TITLE, "target_version": SCRIPT_TARGET_VERSION,
        "max_depth": opts["max_depth"], "max_types": opts["max_types"],
        "samples": opts["samples"], "mask": bool(opts.get("mask", True)), "elapsed": elapsed,
    }
    md = MarkdownWriter(walker, meta, env, globals_, root_status).render()
    md_path = resolve_output_path(opts["output"], opts["label"], env)
    write_text(md_path, md)
    log(tr("log_md", md_path))

    json_path = None
    if opts.get("json"):
        env_json = dict(env)
        env_json["build_type_text"] = build_type_text(env["build_type"])
        data = {
            "meta": meta, "environment": env_json,
            "roots": [{"root": r, "status": s, "message": m} for r, s, m in root_status],
            "globals": {"functions": globals_[0], "others": globals_[1]},
            "types": [walker.types[k].to_dict() for k in walker.order],
        }
        json_path = os.path.splitext(md_path)[0] + ".json"
        write_text(json_path, json.dumps(data, ensure_ascii=False, indent=1, default=to_text))
        log(tr("log_json", json_path))

    stats = {"types": len(walker.order),
             "methods": sum(len(walker.types[k].methods) for k in walker.order),
             "properties": sum(len(walker.types[k].properties) for k in walker.order),
             "cancelled": walker.cancel, "elapsed": elapsed}
    log(tr("log_done", stats["types"], stats["methods"], stats["properties"]))
    return md_path, json_path, stats


def default_opts(env=None):
    env = env or detect_environment()
    label = env.get("default_label") or SCRIPT_TARGET_VERSION
    return {
        "label": label, "env": env,
        "output": os.path.join(default_folder(), default_file_name(label, env)),
        "roots": [r for r, on in ROOT_KEYS if on],
        "max_depth": DEFAULT_MAX_DEPTH, "max_types": DEFAULT_MAX_TYPES, "samples": DEFAULT_SAMPLES,
        "skip": set(DEFAULT_SKIP), "json": True, "mask": True,
    }


# ---------------------------------------------------------------------------
# GUI (System.Windows.Forms)
# ---------------------------------------------------------------------------
def run_gui():
    import clr
    clr.AddReference("System.Windows.Forms")
    clr.AddReference("System.Drawing")
    from System import Decimal
    from System.Diagnostics import Process
    from System.Drawing import Point, Size, Font, Color
    from System.Windows.Forms import (
        Application, Form, Label, TextBox, Button, CheckBox, CheckedListBox,
        NumericUpDown, ProgressBar, SaveFileDialog, DialogResult, MessageBox,
        MessageBoxButtons, MessageBoxIcon, ScrollBars, FormBorderStyle, FormStartPosition)

    opts0 = default_opts()
    env = opts0["env"]
    state = {"running": False, "holder": {}, "auto_path": opts0["output"]}
    W = 700

    form = Form()
    form.Text = SCRIPT_TITLE
    form.ClientSize = Size(W, 760)
    form.FormBorderStyle = FormBorderStyle.FixedDialog
    form.MaximizeBox = False
    form.StartPosition = FormStartPosition.CenterScreen
    try:
        form.Font = Font("Malgun Gothic" if LANG == "ko" else "Segoe UI", 9.0)
    except Exception:
        pass

    def add(ctrl, x, y, w, h):
        ctrl.Location = Point(x, y)
        ctrl.Size = Size(w, h)
        form.Controls.Add(ctrl)
        return ctrl

    def label(text, x, y, w=150, h=20, gray=False):
        l = Label()
        l.Text = text
        l.AutoEllipsis = False
        if gray:
            l.ForeColor = Color.DimGray
        return add(l, x, y, w, h)

    X2 = 120                     # input column
    RW = W - X2 - 12             # input width

    title = label(tr("gui_title", SCRIPT_TARGET_VERSION), 12, 10, W - 24, 24)
    try:
        title.Font = Font(form.Font.FontFamily, 11.0)
    except Exception:
        pass
    label(tr("gui_sub"), 12, 36, W - 24, 20, gray=True)

    # detected environment
    label(tr("gui_detected"), 12, 68, X2 - 16)
    tb_env = add(TextBox(), X2, 65, RW, 24)
    tb_env.ReadOnly = True
    tb_env.Text = env_summary(env) + (
        "" if env.get("target_match") in (True, None) else "  ⚠ " + tr("r_mismatch", SCRIPT_TARGET_VERSION))

    # version label
    label(tr("gui_label"), 12, 102, X2 - 16)
    tb_label = add(TextBox(), X2, 99, 160, 24)
    tb_label.Text = opts0["label"]
    label(tr("gui_label_hint"), X2 + 170, 102, RW - 170, 20, gray=True)

    # save path (Save As)
    label(tr("gui_save"), 12, 138, X2 - 16)
    tb_out = add(TextBox(), X2, 135, RW - 130, 24)
    tb_out.Text = opts0["output"]
    btn_browse = add(Button(), W - 12 - 124, 134, 124, 26)
    btn_browse.Text = tr("gui_browse")
    label(tr("gui_save_hint"), X2, 163, RW, 20, gray=True)

    # roots
    label(tr("gui_roots"), 12, 196, 200)
    clb = add(CheckedListBox(), 12, 218, 200, 190)
    clb.CheckOnClick = True
    for i, (r, on) in enumerate(ROOT_KEYS):
        clb.Items.Add(r)
        clb.SetItemChecked(i, on)

    # options column
    OX = 230
    OW = W - OX - 12
    label(tr("gui_depth"), OX, 199, 150)
    nud_depth = add(NumericUpDown(), OX + 160, 196, 80, 24)
    nud_depth.Minimum, nud_depth.Maximum = Decimal(1), Decimal(40)
    nud_depth.Value = Decimal(DEFAULT_MAX_DEPTH)

    label(tr("gui_types"), OX, 231, 150)
    nud_types = add(NumericUpDown(), OX + 160, 228, 80, 24)
    nud_types.Minimum, nud_types.Maximum = Decimal(10), Decimal(20000)
    nud_types.Value = Decimal(DEFAULT_MAX_TYPES)

    label(tr("gui_samples"), OX, 263, 150)
    nud_samples = add(NumericUpDown(), OX + 160, 260, 80, 24)
    nud_samples.Minimum, nud_samples.Maximum = Decimal(1), Decimal(20)
    nud_samples.Value = Decimal(DEFAULT_SAMPLES)
    label(tr("gui_samples_hint"), OX, 287, OW, 20, gray=True)   # own line, full width

    cb_json = add(CheckBox(), OX, 312, OW, 24)
    cb_json.Text = tr("gui_json")
    cb_json.Checked = True
    cb_mask = add(CheckBox(), OX, 338, OW, 24)
    cb_mask.Text = tr("gui_mask")
    cb_mask.Checked = True

    label(tr("gui_skip"), OX, 366, OW)
    tb_skip = add(TextBox(), OX, 386, OW, 24)
    tb_skip.Text = ", ".join(DEFAULT_SKIP)

    btn_run = add(Button(), 12, 426, 160, 34)
    btn_run.Text = tr("gui_run")
    btn_stop = add(Button(), 180, 426, 110, 34)
    btn_stop.Text = tr("gui_stop")
    btn_stop.Enabled = False
    btn_close = add(Button(), W - 12 - 110, 426, 110, 34)
    btn_close.Text = tr("gui_close")

    prog = add(ProgressBar(), 12, 470, W - 24, 18)
    prog.Minimum, prog.Maximum = 0, 1000
    lbl_status = label(tr("gui_idle"), 12, 492, W - 24)

    tb_log = add(TextBox(), 12, 516, W - 24, 232)
    tb_log.Multiline = True
    tb_log.ReadOnly = True
    tb_log.ScrollBars = ScrollBars.Both
    tb_log.WordWrap = False

    tick = {"n": 0}

    def log(msg):
        tb_log.AppendText(to_text(msg) + "\r\n")
        Application.DoEvents()

    def progress(i, total, ti):
        tick["n"] += 1
        if total:
            prog.Value = min(1000, int(1000.0 * i / total))
        w = state["holder"].get("walker")
        lbl_status.Text = tr("gui_progress", i, total, len(w.types) if w else 0,
                             short(normalize_path(ti.first_expr), 60))
        if tick["n"] % 25 == 0:
            log("  ... %s" % normalize_path(ti.first_expr))
        Application.DoEvents()

    def on_label_changed(sender, e):
        # keep the default file name in sync until the user picks a path himself
        if tb_out.Text == state["auto_path"]:
            folder = os.path.dirname(tb_out.Text) or default_folder()
            new = os.path.join(folder, default_file_name(tb_label.Text.strip() or
                                                         SCRIPT_TARGET_VERSION, env))
            state["auto_path"] = new
            tb_out.Text = new

    def on_browse(sender, e):
        dlg = SaveFileDialog()
        dlg.Title = tr("gui_dlg_save")
        dlg.Filter = "Markdown (*.md)|*.md|All files (*.*)|*.*"
        dlg.DefaultExt = "md"
        dlg.AddExtension = True
        dlg.OverwritePrompt = True
        cur = os.path.expandvars(os.path.expanduser(tb_out.Text.strip().strip('"')))
        folder = cur if os.path.isdir(cur) else os.path.dirname(cur)
        if folder and os.path.isdir(folder):
            dlg.InitialDirectory = folder
        name = os.path.basename(cur) if not os.path.isdir(cur) else ""
        dlg.FileName = name or default_file_name(tb_label.Text.strip() or SCRIPT_TARGET_VERSION, env)
        if dlg.ShowDialog() == DialogResult.OK:
            tb_out.Text = dlg.FileName
            state["confirmed"] = dlg.FileName    # dialog already asked about overwriting

    def on_stop(sender, e):
        w = state["holder"].get("walker")
        if w is not None:
            w.cancel = True
            log(tr("log_stop"))

    def on_close(sender, e):
        on_stop(sender, e)
        form.Close()

    def on_run(sender, e):
        if state["running"]:
            return
        lbl_text = tb_label.Text.strip() or SCRIPT_TARGET_VERSION
        if not tb_out.Text.strip():
            MessageBox.Show(tr("gui_need_path"), SCRIPT_TITLE)
            return
        try:
            out = resolve_output_path(tb_out.Text, lbl_text, env)
        except Exception as ex:
            MessageBox.Show(tr("gui_bad_path", err_text(ex)), SCRIPT_TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        if os.path.exists(out) and state.get("confirmed") != out:
            if MessageBox.Show(tr("gui_overwrite", out), SCRIPT_TITLE, MessageBoxButtons.YesNo,
                               MessageBoxIcon.Question) != DialogResult.Yes:
                return
        tb_out.Text = out
        roots = [to_text(clb.Items[i]) for i in range(clb.Items.Count) if clb.GetItemChecked(i)]
        if not roots:
            MessageBox.Show(tr("gui_need_root"), SCRIPT_TITLE)
            return
        opts = dict(opts0)
        opts.update({
            "output": out, "label": lbl_text, "roots": roots,
            "max_depth": Decimal.ToInt32(nud_depth.Value),
            "max_types": Decimal.ToInt32(nud_types.Value),
            "samples": Decimal.ToInt32(nud_samples.Value),
            "skip": set(s.strip() for s in tb_skip.Text.split(",") if s.strip()),
            "json": bool(cb_json.Checked), "mask": bool(cb_mask.Checked),
        })
        state["running"] = True
        state["holder"] = {}
        btn_run.Enabled = False
        btn_stop.Enabled = True
        tb_log.Clear()
        prog.Value = 0
        try:
            md_path, _json_path, stats = run_dump(opts, log=log, progress=progress,
                                                  walker_holder=state["holder"])
            prog.Value = 1000
            lbl_status.Text = tr("gui_done", stats["elapsed"])
            state["confirmed"] = md_path
            msg = tr("gui_saved", tr("gui_partial") if stats["cancelled"] else "", md_path,
                     stats["types"], stats["methods"], stats["properties"])
            if MessageBox.Show(msg, SCRIPT_TITLE, MessageBoxButtons.YesNo,
                               MessageBoxIcon.Information) == DialogResult.Yes:
                try:
                    Process.Start("explorer.exe", '/select,"%s"' % md_path)
                except Exception:
                    pass
        except Exception as ex:
            log(tr("log_err", err_text(ex)))
            log(traceback.format_exc())
            lbl_status.Text = tr("gui_error")
            MessageBox.Show(tr("gui_fail", err_text(ex)), SCRIPT_TITLE,
                            MessageBoxButtons.OK, MessageBoxIcon.Error)
        finally:
            state["running"] = False
            btn_run.Enabled = True
            btn_stop.Enabled = False

    tb_label.TextChanged += on_label_changed
    btn_browse.Click += on_browse
    btn_run.Click += on_run
    btn_stop.Click += on_stop
    btn_close.Click += on_close

    Application.EnableVisualStyles()
    form.ShowDialog()


def main():
    if not os.environ.get("RS_API_EXPLORER_HEADLESS"):
        try:
            run_gui()
            return
        except ImportError:
            print("clr (GUI) not available - running headless.")
    opts = default_opts()
    if os.environ.get("RS_API_EXPLORER_OUT"):
        opts["output"] = os.environ["RS_API_EXPLORER_OUT"]
    run_dump(opts)


if __name__ == "__main__":
    main()
