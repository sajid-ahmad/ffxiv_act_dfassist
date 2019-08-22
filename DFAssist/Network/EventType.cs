﻿namespace DFAssist
{
    public enum EventType
    {
        // ReSharper disable InconsistentNaming
        NONE,
        INSTANCE_ENTER, // [0] = instance code
        INSTANCE_EXIT,  // [0] = instance code

        // User requests matching
        MATCH_BEGIN,    // [0] = match type(0,1), [1] = roulette code or instance count, [...] = instance

        // Matching state
        MATCH_PROGRESS, // [0] = instance code, [1] = status, [2] = tank, [3] = healer, [4] = dps

        MATCH_ORDER_PROGRESS, // [0] = roulette code, [1] = role order

        // Matched
        MATCH_ALERT,    // [0] = roulette code, [1] = instance code

        // Match end
        MATCH_END,      // [0] = end reason <MatchEndType>
        // ReSharper restore InconsistentNaming
    }
}