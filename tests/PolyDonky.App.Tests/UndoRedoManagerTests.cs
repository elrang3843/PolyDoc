using PolyDonky.App.Services;
using PolyDonky.Core;

namespace PolyDonky.App.Tests;

public class UndoRedoManagerTests
{
    [Fact]
    public void NewManager_HasNoUndoOrRedo()
    {
        var mgr = new UndoRedoManager();
        Assert.False(mgr.CanUndo);
        Assert.False(mgr.CanRedo);
    }

    [Fact]
    public void PushUndo_EnablesUndo()
    {
        var mgr = new UndoRedoManager();
        mgr.PushUndo(MakeDoc("v0"));
        Assert.True(mgr.CanUndo);
        Assert.False(mgr.CanRedo);
    }

    [Fact]
    public void PushUndo_ClearsRedoStack()
    {
        var mgr = new UndoRedoManager();
        mgr.PushUndo(MakeDoc("v0"));
        var redo = mgr.Undo(MakeDoc("v1"));
        Assert.NotNull(redo);
        Assert.True(mgr.CanRedo);

        // 새 PushUndo 가 redo 스택을 비워야 함
        mgr.PushUndo(MakeDoc("v2"));
        Assert.False(mgr.CanRedo);
    }

    [Fact]
    public void Undo_RestoresPreviousSnapshot()
    {
        var mgr  = new UndoRedoManager();
        var v0   = MakeDoc("v0");
        var v1   = MakeDoc("v1");

        mgr.PushUndo(v0);
        var restored = mgr.Undo(v1);

        Assert.NotNull(restored);
        Assert.Equal("v0", FirstParaText(restored!));
    }

    [Fact]
    public void Undo_PushesCurrentToRedo()
    {
        var mgr = new UndoRedoManager();
        mgr.PushUndo(MakeDoc("v0"));
        mgr.Undo(MakeDoc("v1"));

        Assert.True(mgr.CanRedo);
        Assert.False(mgr.CanUndo);
    }

    [Fact]
    public void Redo_RestoresNextSnapshot()
    {
        var mgr = new UndoRedoManager();
        mgr.PushUndo(MakeDoc("v0"));
        mgr.Undo(MakeDoc("v1"));
        var redone = mgr.Redo(MakeDoc("v0"));

        Assert.NotNull(redone);
        Assert.Equal("v1", FirstParaText(redone!));
    }

    [Fact]
    public void UndoRedoCycle_RoundTripsState()
    {
        var mgr = new UndoRedoManager();
        var v0  = MakeDoc("v0");
        var v1  = MakeDoc("v1");
        var v2  = MakeDoc("v2");

        mgr.PushUndo(v0);
        mgr.PushUndo(v1);
        // 현재 상태(외부)는 v2

        var afterUndo1 = mgr.Undo(v2);
        Assert.Equal("v1", FirstParaText(afterUndo1!));

        var afterUndo2 = mgr.Undo(afterUndo1!);
        Assert.Equal("v0", FirstParaText(afterUndo2!));

        Assert.False(mgr.CanUndo);
        Assert.True(mgr.CanRedo);

        var afterRedo1 = mgr.Redo(afterUndo2!);
        Assert.Equal("v1", FirstParaText(afterRedo1!));

        var afterRedo2 = mgr.Redo(afterRedo1!);
        Assert.Equal("v2", FirstParaText(afterRedo2!));
    }

    [Fact]
    public void Undo_ReturnsNull_WhenStackEmpty()
    {
        var mgr = new UndoRedoManager();
        Assert.Null(mgr.Undo(MakeDoc("v0")));
    }

    [Fact]
    public void Redo_ReturnsNull_WhenStackEmpty()
    {
        var mgr = new UndoRedoManager();
        Assert.Null(mgr.Redo(MakeDoc("v0")));
    }

    [Fact]
    public void Clear_EmptiesBothStacks()
    {
        var mgr = new UndoRedoManager();
        mgr.PushUndo(MakeDoc("v0"));
        mgr.Undo(MakeDoc("v1"));

        Assert.True(mgr.CanRedo);
        mgr.Clear();
        Assert.False(mgr.CanUndo);
        Assert.False(mgr.CanRedo);
    }

    [Fact]
    public void StateChanged_FiresOnPushUndo()
    {
        var mgr  = new UndoRedoManager();
        int hits = 0;
        mgr.StateChanged += (_, _) => hits++;

        mgr.PushUndo(MakeDoc("v0"));
        Assert.Equal(1, hits);
    }

    [Fact]
    public void Snapshot_IsDeepCopy_NotSameReference()
    {
        var mgr = new UndoRedoManager();
        var v0  = MakeDoc("original");
        mgr.PushUndo(v0);

        var restored = mgr.Undo(MakeDoc("after"));

        Assert.NotNull(restored);
        Assert.NotSame(v0, restored);
        // 외부에서 v0 을 변형해도 복원본에는 영향이 없어야 함
        ((Paragraph)v0.Sections[0].Blocks[0]).Runs[0].Text = "mutated";
        Assert.Equal("original", FirstParaText(restored!));
    }

    [Fact]
    public void MaxDepth_LimitsUndoStack()
    {
        var mgr = new UndoRedoManager();
        // 한도 + 5 만큼 push 했을 때 정확히 한도 만큼만 남는지.
        for (int i = 0; i < UndoRedoManager.MaxDepth + 5; i++)
            mgr.PushUndo(MakeDoc($"v{i}"));

        // 끝까지 Undo 했을 때 가장 오래된 항목은 폐기되었으므로
        // 첫 Undo 가 반환하는 값은 v(MaxDepth + 4) ... 가 아닌 가장 마지막 push 값.
        // 깊이만 확인 — 정확히 MaxDepth 번 undo 가능.
        int undoCount = 0;
        var current   = MakeDoc("current");
        while (mgr.CanUndo)
        {
            current = mgr.Undo(current)!;
            undoCount++;
        }
        Assert.Equal(UndoRedoManager.MaxDepth, undoCount);
    }

    private static PolyDonkyument MakeDoc(string firstParaText)
    {
        var doc     = new PolyDonkyument();
        var section = new Section();
        var p       = new Paragraph();
        p.AddText(firstParaText);
        section.Blocks.Add(p);
        doc.Sections.Add(section);
        return doc;
    }

    private static string FirstParaText(PolyDonkyument doc)
    {
        var p = (Paragraph)doc.Sections[0].Blocks[0];
        return p.Runs[0].Text;
    }
}
