"""Register one main pipeline (steps 0..5) per entity, plus separate side pipelines.

  kedro run --pipeline=galaxy            # main 6-step pipeline for Galaxy
  kedro run --pipeline=galaxy_prebuild   # source extraction (entities with prebuild_sources only)
  kedro run --pipeline=galaxy_export     # tagged-rows xlsx export
  kedro run --pipeline=galaxy_summary    # re-print step 3 summary stats
  kedro run --pipeline=galaxy_validate   # structural LLM tagging quality checks
  kedro run --pipeline=galaxy_offsets    # offset/reversal detection for audit sampling
  kedro run --pipeline=all               # all entities' main pipelines
  kedro run --pipeline=prebuild          # all entities that have prebuild_sources
  kedro run --pipeline=audit             # CROSS-COMPANY audit sample (one file covering all 6)
  kedro run --pipeline=export            # all entities' exports
  kedro run --pipeline=summary           # all entities' summaries
  kedro run --pipeline=validate          # all entities' validations
  kedro run --pipeline=offsets           # all entities' offset detections
  kedro run --pipeline=table2            # parse 12 表二 forms + LLM-analyze each project
  kedro run --pipeline=table2_extract    # 表二 parse only (no LLM)
  kedro run --pipeline=table2_analyze    # LLM only (requires _extracted.xlsx exists)
"""
import logging

from kedro.pipeline import Pipeline

from kpi.lib.conf import folder_for, load_master
from kpi.pipelines.step0_convert import create_pipeline as p0
from kpi.pipelines.step0_prebuild import create_pipeline as p0_pre
from kpi.pipelines.step0_5_split_year import create_pipeline as p0_5
from kpi.pipelines.step1_extract import create_pipeline as p1
from kpi.pipelines.step2_tag_projects import create_pipeline as p2
from kpi.pipelines.step3_tag_signatures import create_pipeline as p3
from kpi.pipelines.step4_apply_tags import create_pipeline as p4
from kpi.pipelines.step4_5_derive_counts import create_pipeline as p4_5
from kpi.pipelines.step5_build_report import create_pipeline as p5
from kpi.pipelines.step6_master_audit import create_pipeline as p6
from kpi.pipelines.audit import create_pipeline as p_audit
from kpi.pipelines.export_xlsx import create_pipeline as p_export
from kpi.pipelines.show_summary import create_pipeline as p_summary
from kpi.pipelines.validate import create_pipeline as p_validate
from kpi.pipelines.offsets import create_pipeline as p_offsets
from kpi.pipelines.table2 import create_pipeline as p_table2
from kpi.pipelines.table2.pipeline import (
    create_analyze_only_pipeline as p_table2_analyze,
    create_extract_only_pipeline as p_table2_extract,
)
from kpi.pipelines.generate import create_pipeline as p_generate
from kpi.pipelines.tableau_export import create_pipeline as p_tableau

log = logging.getLogger(__name__)


def _make_main_pipeline(entity_key: str) -> Pipeline:
    return (p0(entity_key) + p0_5(entity_key) + p1(entity_key) + p2(entity_key)
            + p3(entity_key) + p4(entity_key) + p4_5(entity_key) + p5(entity_key)
            + p6(entity_key))


def register_pipelines() -> dict[str, Pipeline]:
    pipelines: dict[str, Pipeline] = {}
    main_combined = Pipeline([])
    prebuild_combined = Pipeline([])
    export_combined = Pipeline([])
    summary_combined = Pipeline([])
    validate_combined = Pipeline([])
    offsets_combined = Pipeline([])

    try:
        master = load_master()
    except Exception as _e:
        import traceback as _tb
        log.warning("load_master() failed — no pipelines registered: %s", _e)
        _tb.print_exc()
        master = {"companies": {}}
    companies = master.get("companies") or {}

    for key, ent in companies.items():
        if not ent or not ent.get("name"):
            continue

        try:
            main_p = _make_main_pipeline(key)
            export_p = p_export(key)
            summary_p = p_summary(key)
            validate_p = p_validate(key)
            offsets_p = p_offsets(key)
        except Exception as _e:
            import traceback as _tb
            log.warning("Failed to build pipelines for %s: %s", key, _e)
            _tb.print_exc()
            continue

        pipelines[key] = main_p
        pipelines[f"{key}_export"] = export_p
        pipelines[f"{key}_summary"] = summary_p
        pipelines[f"{key}_validate"] = validate_p
        pipelines[f"{key}_offsets"] = offsets_p

        # Prebuild pipeline — only for entities with prebuild_sources OR prebuild_years
        if ent.get("prebuild_sources") or ent.get("prebuild_years"):
            try:
                pre_p = p0_pre(key)
                pipelines[f"{key}_prebuild"] = pre_p
                prebuild_combined = prebuild_combined + pre_p
                alias = (ent.get("alias") or "").strip()
                if alias and alias != key:
                    pipelines[f"{alias}_prebuild"] = pre_p
            except Exception as _e:
                log.warning("Failed to build prebuild pipeline for %s: %s", key, _e)

        alias = (ent.get("alias") or "").strip()
        if alias and alias != key:
            pipelines[alias] = main_p
            pipelines[f"{alias}_export"] = export_p
            pipelines[f"{alias}_summary"] = summary_p
            pipelines[f"{alias}_validate"] = validate_p
            pipelines[f"{alias}_offsets"] = offsets_p

        main_combined = main_combined + main_p
        export_combined = export_combined + export_p
        summary_combined = summary_combined + summary_p
        validate_combined = validate_combined + validate_p
        offsets_combined = offsets_combined + offsets_p

    pipelines["all"] = main_combined
    pipelines["prebuild"] = prebuild_combined
    pipelines["export"] = export_combined
    pipelines["summary"] = summary_combined
    pipelines["validate"] = validate_combined
    pipelines["offsets"] = offsets_combined

    # Standalone (non-entity) pipelines.
    try:
        pipelines["audit"] = p_audit()
    except Exception as _e:
        log.warning("Failed to register audit pipeline: %s", _e)
    try:
        pipelines["table2"] = p_table2()
        pipelines["table2_extract"] = p_table2_extract()
        pipelines["table2_analyze"] = p_table2_analyze()
    except Exception as _e:
        log.warning("Failed to register table2 pipelines: %s", _e)
    try:
        pipelines["generate"] = p_generate()   # delivery bundle: 投資方向 (data/review) + tableau (data/tableau)
    except Exception as _e:
        log.warning("Failed to register generate pipeline: %s", _e)
    try:
        pipelines["tableau"] = p_tableau()      # delivery: tableau only (data/tableau)
    except Exception as _e:
        log.warning("Failed to register tableau pipeline: %s", _e)

    pipelines["__default__"] = next(iter(pipelines.values()), Pipeline([]))
    return pipelines
