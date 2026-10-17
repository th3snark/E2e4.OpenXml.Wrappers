using Xunit;

// The tests share fixture documents on disk (some are opened editable), so classes must not run in parallel.
[assembly: CollectionBehavior(DisableTestParallelization = true)]
