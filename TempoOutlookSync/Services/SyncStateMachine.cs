namespace TempoOutlookSync.Services;

using TempoOutlookSync.Common;

public sealed class SyncStateMachine
{
    private readonly Lock _lock;

    private SyncState currentState;
    private SyncBlocker blockers;

    public SyncState State
    {
        get
        {
            using (_ = _lock.EnterScope())
            {
                return currentState;
            }
        }
    }

    public SyncBlocker Blockers
    {
        get
        {
            using (_ = _lock.EnterScope())
            {
                return blockers;
            }
        }
    }

    public bool IsSyncing
    {
        get
        {
            using ( _ = _lock.EnterScope())
            {
                return currentState is SyncState.Syncing;
            }
        }
    }

    public bool IsBlocked
    {
        get
        {
            using (_ = _lock.EnterScope())
            {
                return blockers is not SyncBlocker.None;
            }
        }
    }

    public event Action? StateChanged;

    public SyncStateMachine()
    {
        _lock = new Lock();
        currentState = SyncState.Idle;
        blockers = SyncBlocker.None;
    }

    public bool TryStartSync()
    {
        using (_ = _lock.EnterScope())
        {
            if (currentState is not SyncState.Idle || blockers is not SyncBlocker.None) return false;

            currentState = SyncState.Syncing;
        }

        StateChanged?.Invoke();
        return true;
    }

    public void FinishSync()
    {
        using (_ = _lock.EnterScope())
        {
            currentState = SyncState.Idle;
        }

        StateChanged?.Invoke();
    }

    public void AddBlocker(SyncBlocker blocker)
    {
        using (_ = _lock.EnterScope())
        {
            blockers |= blocker;
        }

        StateChanged?.Invoke();
    }

    public void RemoveBlocker(SyncBlocker blocker)
    {
        using (_ = _lock.EnterScope())
        {
            blockers &= ~blocker;
        }

        StateChanged?.Invoke();
    }
}