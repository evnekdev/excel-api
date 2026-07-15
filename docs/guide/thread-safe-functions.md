# Thread-safe functions

Mark a generated UDF `thread_safe` only when it injects `ThreadSafeContext` (or
needs no context) and its entire implementation is safe under Excel's
multi-threaded calculation contract. It must not call a worksheet-, macro-, or
lifecycle-only Excel API, mutate global non-thread-safe state, or wait for
another Excel callback.

Thread-safe is a registration and implementation promise. Test pure logic away
from Excel and validate the resulting XLL under multi-threaded recalculation on
a supported host.
