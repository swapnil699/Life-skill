**Load Balancers:**

* Load balancers distribute incoming network traffic across multiple
servers to ensure that no single server is overwhelmed.

![Load Balancer](https://substackcdn.com/image/fetch/$s_!QKkY!,w_1456,c_limit,f_webp,q_auto:good,fl_progressive:steep/https%3A%2F%2Fsubstack-post-media.s3.amazonaws.com%2Fpublic%2Fimages%2Fa5215d10-fcd8-4c50-acb6-ccbb36dea24f_832x1000.png)

**Popular Load Balancer Algorithms:**

1.Round Robin

2.Weighted Round Robin

3.Least connections

 **How does LB know if the server is available or not?** 
 
 **Heart check:**

* Your Application Load Balancer periodically sends requests to its
registered targets to test their status. These tests are called health
checks.

* The load balancer keeps pinging the server every x(sec) time.

* If a server doesn\'t respond after 3 consecutive attempts, the load
balancer assumes the server is dead.

**Heart Break:**

* The server is expected to send a request to the load balancer every x
time.

* If the load balancer doesn\'t get any request from the server for 3\*x
seconds, the load balancer assumes you are dead and removes it.

**How a Load Balancer forwards requests to the server:**


**Round Robin Algorithms:**

**How it Works:**

* A request is sent to the first server in the list.

* The next request is sent to the second server, and this process
continues.

* After the last server in the list, the algorithm loops back to the first
server.

**When to Use:**

* When all servers have similar processing capabilities and are equally
capable of handling requests.

* When simplicity and even distribution of load are more critical.

**Benefits:**

* Simple to implement and understand.

* Ensures even distribution of traffic.

**Drawbacks:**

* Does not consider server load or response time.

* Availability and Latency :

**Availability:**

* Availability is a measure of the probability that a particular server or
service is operational at any given time. It is typically expressed as a
percentage.

* Formula: Availability = uptime / (uptime + downtime)

**Latency:**

* Latency refers to the time it takes for a particular operation to
complete.

* Low latency means faster response time and a more responsive system.

* High latency can lead to lag and poor user experience.

**Database scaling:**

There are two broad approaches to database scaling: vertical scaling and
horizontal scaling.
![Vertical scaling and
Horizontal scaling](https://media.geeksforgeeks.org/wp-content/uploads/20250807104547934960/vertical_and_horizontal_scaling.webp)

**Vertical scaling:**

* Involves boosting the power of an existing machine within your system to
handle increased loads.


* This can mean upgrading the CPU, RAM, storage, and other hardware
components to boost the server\'s capacity.

\* Upgrading the CPU: Replacing your server\'s processor with a more
powerful one. \* Increasing RAM: Adding more memory to handle larger
datasets. \* Enhancing Storage: Switching to faster storage (SSDs) or
increasing overall storage capacity.

**Pros of Vertical Scaling:**

1.Simple setup

2.Lower latency

**Cons of Vertical Scaling:**

1.Single point of failure

2.There is a practical limit -\> can\'t scale infinitely

3.Cost increase

4.Updating application downtime

**Horizontal scaling:**

* Horizontal scaling involves adding more servers.

* Each server runs a copy of the application, and the load is balanced
among them, often using a load balancer.


**Pros of Horizontal Scaling**

1.Infinitely scalable

2.Cost increases linearly

3.Improved fault tolerance

4\. A better geographical experience

**Cons of Horizontal Scaling**

1.It makes the complex infrastructure

2.All issues with distributed system data

3.Increased latency

**References Section**

**Load balancer:**

https://bytebytego.com/courses/system-design-interview/scale-from-zero-to-millions-of-users

https://docs.aws.amazon.com/elasticloadbalancing/latest/userguide/what-is-load-balancing.html

https://medium.com/@abhirup.acharya009/load-balancing-system-design-fundamentals-d64674227c36

https://www.baeldung.com/cs/load-balancer

**Algorithm 1: Round Robin:**

https://blog.algomaster.io/p/load-balancing-algorithms-explained-with-code

**Horizontally and Vertically:**

https://www.baeldung.com/cs/scaling-horizontally-vertically

https://medium.com/design-microservices-architecture-with-patterns/scalability-vertical-scaling-horizontal-scaling-adb52ff679f
