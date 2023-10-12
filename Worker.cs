using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using System.Text;

namespace RabbitPracticeProject
{
    public class Worker : BackgroundService
    {
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                Console.WriteLine("oe");
                await Task.Delay(1);
                ConnectionFactory connectionFactory = new ConnectionFactory();
                connectionFactory.Uri = new Uri("amqp://guest:guest@localhost:5672");

                Semaphore sem = new Semaphore(0, 1);
                using (var connection = connectionFactory.CreateConnection())
                {
                    using (var channel = connection.CreateModel())
                    {
                        channel.BasicQos(0, 1, false);


                        var consumer = new EventingBasicConsumer(channel);

                        consumer.Received += (model, message) =>
                        {
                            string messageToString = Encoding.UTF8.GetString(message.Body.ToArray());

                            if (messageToString == "morcha")
                            {
                                channel.BasicAck(deliveryTag: message.DeliveryTag, false);
                                sem.Release();
                                return;
                            }

                            Console.WriteLine(messageToString);
                            channel.BasicAck(deliveryTag: message.DeliveryTag, false);
                        };

                        channel.BasicConsume(queue: "excelQueue", autoAck: false, consumer: consumer);
                        sem.WaitOne();
                        Console.WriteLine("DONE");
                        return;
                    }
                }
            }
            
        }
    }
}
