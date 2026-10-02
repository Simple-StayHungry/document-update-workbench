from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]


def test_fullscreen_shell_has_no_outer_frame():
    css=(ROOT/'web/style.css').read_text(encoding='utf-8')
    assert '.stage{width:100%;height:100vh;margin:0;padding:0}' in css
    assert '.app-shell{width:100%;height:100vh;min-height:0' in css
    assert 'border-radius:26px' not in css


def test_import_and_done_surfaces_share_optical_center():
    css=(ROOT/'web/style.css').read_text(encoding='utf-8')
    html=(ROOT/'web/index.html').read_text(encoding='utf-8')
    assert '.import-center{width:min(900px,100%);display:flex;flex-direction:column;align-items:center;transform:translateY(-4vh)}' in css
    assert '.done-card{width:min(560px,92vw);text-align:center;padding:26px 30px 24px;background:transparent;transform:translateY(-4vh)}' in css
    assert '把文件放进来' in html
    assert '选择文件夹' in html


def test_no_manual_review_skips_detail_workspace_and_goes_to_export():
    js=(ROOT/'web/app.js').read_text(encoding='utf-8')
    html=(ROOT/'web/index.html').read_text(encoding='utf-8')
    assert "if(!manualItems().length){renderNoManual();return}" in js
    assert '无需人工确认' in html
    assert '生成修订版' in html
    assert 'reviewDoneView' in html


def test_manual_review_is_decision_first_compare_surface():
    js=(ROOT/'web/app.js').read_text(encoding='utf-8')
    html=(ROOT/'web/index.html').read_text(encoding='utf-8')
    css=(ROOT/'web/style.css').read_text(encoding='utf-8')
    assert 'manualReviewView' in html
    assert 'manual-workspace' in html
    assert 'decision-compare' in js
    assert '修改章节 ·' in js
    assert '只显示真正需要你决定的事项' in html
    assert '.decision-compare{display:grid;grid-template-columns:minmax(0,1fr) minmax(0,1fr)' in css


def test_manual_update_requires_explicit_source_choice():
    js=(ROOT/'web/app.js').read_text(encoding='utf-8')
    assert "if(decision==='update'&&chosen==null)return" in js
    assert "${chosen==null?'disabled':''}" in js
    assert '先选择来源，或保持原文' in js


def test_export_download_stays_in_center_completion_surface():
    js=(ROOT/'web/app.js').read_text(encoding='utf-8')
    html=(ROOT/'web/index.html').read_text(encoding='utf-8')
    assert 'doneDownloadBtn' in html
    assert "$('#doneDownloadBtn').classList.remove('hidden')" in js
    assert 'id="donePanel"' not in html
