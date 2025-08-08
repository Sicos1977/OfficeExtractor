## About

Provides the HashCode type for .NET Standard 2.0. This package is not required starting with .NET Standard 2.1 and .NET Core 3.0.

The HashCode type is shipped as part of the shared framework starting with .NET 5.

## Key Features

The HashCode type offered in this assembly combines the hash code for multiple values into a single hash code.

## How to Use

The static methods in this class combine the default hash codes of _up to eight_ values.

```cs
using System;
using System.Collections.Generic;

public struct OrderOrderLine : IEquatable<OrderOrderLine>
{
    public int OrderId { get; }
    public int OrderLineId { get; }

    public OrderOrderLine(int orderId, int orderLineId) => (OrderId, OrderLineId) = (orderId, orderLineId);

    public override bool Equals(object obj) => obj is OrderOrderLine o && Equals(o);

    public bool Equals(OrderOrderLine other) => OrderId == other.OrderId && OrderLineId == other.OrderLineId;

    public override int GetHashCode() => HashCode.Combine(OrderId, OrderLineId);
}

class Program
{
    static void Main(string[] args)
    {
        var set = new HashSet<OrderOrderLine>
        {
            new OrderOrderLine(1, 1),
            new OrderOrderLine(1, 1),
            new OrderOrderLine(1, 2)
        };

        Console.WriteLine($"Item count: {set.Count}.");
    }
}
// The example displays the following output:
// Item count: 2.
```

The instance methods in this class combine the hash codes of _more than eight_ values.

```cs
using System;
using System.Collections.Generic;

public struct Path : IEquatable<Path>
{
    public IReadOnlyList<string> Segments { get; }

    public Path(params string[] segments) => Segments = segments;

    public override bool Equals(object obj) => obj is Path o && Equals(o);

    public bool Equals(Path other)
    {
        if (ReferenceEquals(Segments, other.Segments)) return true;
        if (Segments is null || other.Segments is null) return false;
        if (Segments.Count != other.Segments.Count) return false;

        for (var i = 0; i < Segments.Count; i++)
        {
            if (!string.Equals(Segments[i], other.Segments[i]))
                return false;
        }

        return true;
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();

        for (var i = 0; i < Segments?.Count; i++)
            hash.Add(Segments[i]);

        return hash.ToHashCode();
    }
}

class Program
{
    static void Main(string[] args)
    {
        var set = new HashSet<Path>
        {
            new Path("C:", "tmp", "file.txt"),
            new Path("C:", "tmp", "file.txt"),
            new Path("C:", "tmp", "file.tmp")
        };

        Console.WriteLine($"Item count: {set.Count}.");
    }
}
// The example displays the following output:
// Item count: 2.
```

The instance methods also combine the hash codes produced by a specific IEqualityComparer<T> implementation.

```cs
using System;
using System.Collections.Generic;

public struct Path : IEquatable<Path>
{
    public IReadOnlyList<string> Segments { get; }

    public Path(params string[] segments) => Segments = segments;

    public override bool Equals(object obj) => obj is Path o && Equals(o);

    public bool Equals(Path other)
    {
        if (ReferenceEquals(Segments, other.Segments)) return true;
        if (Segments is null || other.Segments is null) return false;
        if (Segments.Count != other.Segments.Count) return false;

        for (var i = 0; i < Segments.Count; i++)
        {
            if (!string.Equals(Segments[i], other.Segments[i], StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();

        for (var i = 0; i < Segments?.Count; i++)
            hash.Add(Segments[i], StringComparer.OrdinalIgnoreCase);

        return hash.ToHashCode();
    }
}

class Program
{
    static void Main(string[] args)
    {
        var set = new HashSet<Path>
        {
            new Path("C:", "tmp", "file.txt"),
            new Path("C:", "TMP", "file.txt"),
            new Path("C:", "tmp", "FILE.TXT")
        };

        Console.WriteLine($"Item count: {set.Count}.");
    }
}
// The example displays the following output:
// Item count: 1.
```

The HashCode structure must be passed by-reference to other methods, as it is a value type.
```cs
using System;
using System.Collections.Generic;

public struct Path : IEquatable<Path>
{
    public IReadOnlyList<string> Segments { get; }

    public Path(params string[] segments) => Segments = segments;

    public override bool Equals(object obj) => obj is Path o && Equals(o);

    public bool Equals(Path other)
    {
        if (ReferenceEquals(Segments, other.Segments)) return true;
        if (Segments is null || other.Segments is null) return false;
        if (Segments.Count != other.Segments.Count) return false;

        for (var i = 0; i < Segments.Count; i++)
        {
            if (!PlatformUtils.PathEquals(Segments[i], other.Segments[i]))
                return false;
        }

        return true;
    }

    public override int GetHashCode()
    {
        var hash = new HashCode();

        for (var i = 0; i < Segments?.Count; i++)
            PlatformUtils.AddPath(ref hash, Segments[i]);

        return hash.ToHashCode();
    }
}

internal static class PlatformUtils
{
    public static bool PathEquals(string a, string b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    public static void AddPath(ref HashCode hash, string path) => hash.Add(path, StringComparer.OrdinalIgnoreCase);
}

class Program
{
    static void Main(string[] args)
    {
        var set = new HashSet<Path>
        {
            new Path("C:", "tmp", "file.txt"),
            new Path("C:", "TMP", "file.txt"),
            new Path("C:", "tmp", "FILE.TXT")
        };

        Console.WriteLine($"Item count: {set.Count}.");
    }
}
// The example displays the following output:
// Item count: 1.
```

## Main Types

The main types provided by this library are:

- System.HashCode

## Additional Documentation

- [Security design doc for System.HashCode](https://github.com/dotnet/runtime/tree/main/docs/design/security/System.HashCode.md)
- API reference can be found in: https://learn.microsoft.com/en-us/dotnet/api/system.hashcode

## License

Microsoft.Bcl.HashCode is released as open source under the [MIT license](https://licenses.nuget.org/MIT).