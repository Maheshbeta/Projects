import psutil
import time
import matplotlib.pyplot as plt

def track_data_usage():
    data_usage = []
    timestamps = []

    # Initial network stats
    initial_stats = psutil.net_io_counters()
    start_time = time.time()

    while True:
        # Current network stats
        current_stats = psutil.net_io_counters()
        elapsed_time = time.time() - start_time

        # Calculate data usage since the last check
        upload = current_stats.bytes_sent - initial_stats.bytes_sent
        download = current_stats.bytes_recv - initial_stats.bytes_recv

        # Append data usage and timestamp to the lists
        data_usage.append((upload, download))
        timestamps.append(elapsed_time)

        # Update initial statistics for the next iteration
        initial_stats = current_stats

        # Check if the user wants to stop tracking
        stop_tracking = input("Press 'q' to stop tracking or any other key to continue: ")
        if stop_tracking.lower() == 'q':
            break

    # Generate a graph showing the data usage over time
    plot_graph(timestamps, data_usage)


def plot_graph(timestamps, data_usage):
    uploads = [upload for upload, _ in data_usage]
    downloads = [download for _, download in data_usage]

    plt.plot(timestamps, uploads, label='Uploads')
    plt.plot(timestamps, downloads, label='Downloads')

    plt.xlabel('Time (seconds)')
    plt.ylabel('Data Usage (bytes)')
    plt.title('Data Usage During Online Session')
    plt.legend()
    plt.grid(True)
    plt.show()

# Start tracking data usage
track_data_usage()

#Psutil may not function correctly if user does not have access to network data
