using Quartz;
using WACS.Core;

IHost builder = Host.CreateDefaultBuilder(args)
    .ConfigureServices((ctx, services) =>
    {
        services.AddQuartz(q =>
        {
            // q.UseMicrosoftDependencyInjectionJobFactory();
            q.UseSimpleTypeLoader();
            q.UseInMemoryStore();
            q.UseDefaultThreadPool(tp =>
            {
                tp.MaxConcurrency = 10;
            });
        });
        services.AddQuartzHostedService(opt =>
        {
            opt.WaitForJobsToComplete = true;
        });
    })
    .UseWindowsService()
    .Build();

var schedulerFactory = builder.Services.GetRequiredService<ISchedulerFactory>();
var scheduler = await schedulerFactory.GetScheduler();

// define the job and tie it to our HelloJob class
var job = JobBuilder.Create<Job>()
    .WithIdentity("myJob", "group1")
    .Build();

// Trigger the job to run now, and then every 40 seconds
// "0 0 8 1-3 * ?"
var trigger = TriggerBuilder.Create()
    .StartNow()
    // .WithCronSchedule("0 45 19 1-5 * ?")
    .Build();

await scheduler.ScheduleJob(job, trigger);
await builder.RunAsync();

