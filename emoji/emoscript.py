class _Emo:
    __slots__ = ("_map",)
    def __init__(self):
        self._map = {
            "ok": "✅",
            "ng": "❌",
            "warn": "⚠️",
            "info": "ℹ️",
            "star": "⭐",
            "spark": "✨",
            "wait": "⏳",
            "done": "🟢",
            "stop": "⛔",
        }

    def __getattr__(self, name: str) -> str:
        # 定義がなければ :name: で返す（Slack風プレースホルダ）
        return self._map.get(name, f":{name}:")

    def add(self, name: str, value: str) -> None:
        """絵文字を追加/上書き"""
        if not name:
            raise ValueError("name must not be empty")
        self._map[name] = value

    def remove(self, name: str) -> None:
        """絵文字を削除（存在しなくてもOK）"""
        self._map.pop(name, None)

    def __getitem__(self, name: str) -> str:
        # emo["ok"] でも使える
        return getattr(self, name)

# グローバルに1個だけ使う想定
emo = _Emo()

# ============== 実行部 ==============
if __name__ == "__main__":
    
    print(f"{emo.ok} 処理完了")         # ✅ 処理完了
    print(f"{emo.warn} リトライします")  # ⚠️ リトライします
    print(emo.ng)                        # ❌
    print(emo["star"])                   # ⭐
    print(f"{emo.unknown}")              # :unknown:

    # 追加もできる
    emo.add("ship", "🚢")
    print(f"{emo.ship} 出荷済み")        # 🚢 出荷済み
