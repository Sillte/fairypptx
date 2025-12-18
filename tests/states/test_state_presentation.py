from fairypptx.presentation import Presentation
from fairypptx.states.presentation import PresentationStateModel
from fairypptx.core.resolvers import Application

def test_presentation_state_apply_reorder():
    # 準備: 元となるプレゼン(source)と適用先(target)を作成
    app = Application().api
    pres_source = Presentation(app.Presentations.Add())
    pres_target = Presentation(app.Presentations.Add())

    try:
        # 1. Sourceにスライドを追加し、名前(ID)を付けて順番をバラバラにする
        # IDを固定するためにスライドのNameプロパティなどを利用している想定
        _ = pres_source.slides.add()
        _ = pres_source.slides.add()
        _ = pres_source.slides.add()
        
        # モデル抽出 (S1, S2, S3 の順)
        source_model = PresentationStateModel.from_entity(pres_source)

        # 2. Targetにも同じIDを持つスライドを「違う順番」で作る
        # (S3, S1, S2) の順番で作成
        _ = pres_target.slides.add()
        _ = pres_target.slides.add()
        _ = pres_target.slides.add()
        
        # IDを手動で同期（本来はSlideStateModel.from_entityで取得されるID）
        # テスト用に、targetのスライドIDをsourceと同じものに強制書き換え
        for i, s_src in enumerate(pres_source.slides):
            pres_target.slides[i].api.Name = s_src.api.Name
        
        # Targetを逆順に並べ替えておく (初期状態をバラバラにする)
        pres_target.slides.reorder([3, 2, 1])
        
        # 3. Apply実行 (source_model の順番 [S1, S2, S3] を target に適用)
        pres_target = source_model.apply(pres_target)

        # 4. 検証
        assert len(pres_target.slides) == 3
        for i in range(len(pres_target.slides)):
            # TargetのスライドIDが、Sourceモデルの定義通りの順番になっているか
            assert pres_target.slides[i].api.SlideID == source_model.slides[i].id
            # SlideIndexが 1, 2, 3... と正しく並んでいるか
            assert pres_target.slides[i].index == i + 1

    finally:
        # 後片付け
        pres_source.api.Close()
        pres_target.api.Close()


