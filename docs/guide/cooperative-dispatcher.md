# Cooperative dispatcher

**Preview — live pump validation pending.** The dispatcher transports owned,
closed-catalogue operations to a later legal Excel callback. It has bounded
pending capacity and batch size, tickets, generation isolation, capability
matching, cancellation, cooperative expiry, and panic-safe retirement.

**Enqueueing a request does not wake Excel.** A request remains queued until a
user invokes `RUST.DISPATCH.PUMP` or another explicitly approved Excel-issued
callback drains compatible work. There are no timers, hidden windows,
`PostMessage`, COM, RTD, or `xlcOnTime` wake mechanisms in the core dispatcher.
Do not wait on a ticket from an Excel callback thread.
