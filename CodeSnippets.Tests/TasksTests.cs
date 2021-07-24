using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace CodeSnippets.Tests
{
    public class TasksTests
    {
        private readonly ITestOutputHelper _output;

        public TasksTests(ITestOutputHelper output)
        {
            _output = output ?? throw new ArgumentNullException(nameof(output));
        }

        [Fact]
        public async Task CanCreateAndRunTasks()
        {
            var tasks = new List<Task>
            {
                new Task(() => _output.WriteLine("Task #1")),
                new Task(() => _output.WriteLine("Task #2"))
            };

            tasks.ForEach(t => t.Start());

            await Task.WhenAll(tasks);
        }
    }
}
