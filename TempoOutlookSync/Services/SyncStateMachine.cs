namespace TempoOutlookSync.Services;

using TempoOutlookSync.Common;

public sealed class SyncStateMachine
{
    private readonly Lock _lock;

    private SyncState currentState;

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

    public bool CanSync => State is SyncState.Idle;
    public bool IsSyncing => State is SyncState.Syncing;

    public event Action<SyncState>? StateChanged;

    public SyncStateMachine()
    {
        _lock = new Lock();
        currentState = SyncState.Idle;
    }

    public bool Transition(SyncState state)
    {
        using (_ = _lock.EnterScope())
        {
            if (currentState == state) return false;

            if (!IsValidTransition(currentState, state)) return false;

            currentState = state;
            StateChanged?.Invoke(currentState);
            return true;
        }
    }

    private static bool IsValidTransition(SyncState from, SyncState to)
    {
        return from switch
        {
            SyncState.Idle => to is SyncState.Syncing or SyncState.Blocked,
            SyncState.Syncing => to is SyncState.Idle or SyncState.Blocked,
            SyncState.Blocked => to is SyncState.Idle,
            _ => false
        };
    }
}