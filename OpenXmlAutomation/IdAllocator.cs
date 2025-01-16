namespace OpenXmlAutomation;

/// <summary>
/// Allocator for unique IDs smaller than GUIDs
/// </summary>

public class IdAllocator
{
    int id32 = 0;

    /// <summary>
    /// Create an ID allocator, initialising it with the
    /// number of seconds since the 2025 epoch.
    /// </summary>

    public IdAllocator()
    {
        NextId();
    }

    /// <summary>
    /// Create an ID allocator with a given seed value
    /// </summary>
    /// <param name="seed">The seed value to use</param>
    /// <exception cref="ArgumentException">Thrown if
    /// the seed was zero or negative</exception>

    public IdAllocator(int seed)
    {
        if (seed <= 0)
            throw new ArgumentException
                ("Id allocator seed must be positive and non-zero");
        id32 = seed;
    }

    private readonly static DateTime epoch = new
            (2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);

    /// <summary>
    /// Allocate a unique Id based on current date and time. Wraps
    /// at epoch + 68 years (plus an hour and a bit).
    /// </summary>
    /// <returns>The next sequential ID value as a positive integer</returns>

    public int NextId()
    {
        int since2k = int.MaxValue & (int)((DateTime.UtcNow - epoch).TotalSeconds);
        if (since2k > id32)
            id32 = since2k;
        else
            id32++;
        return id32;
    }

    public int NextRandomId()
    {
        // 31-bit primitive binary polynomial
        // to generate a sequence with 2^31 - 1
        // repetition period

        const int ibp = 0x4BC915C3;

        bool doXor = (id32 & 1) != 0;
        id32 >>= 1;
        if (doXor)
            id32 ^= ibp;
        return id32;
    }
}
