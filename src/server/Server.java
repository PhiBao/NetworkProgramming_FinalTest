package server;

import java.io.IOException;
import java.net.*;

public class Server {

	@SuppressWarnings("resource")
	public static void main(String args[]) {
		try {
			// Tao cong 8039 de server lang nghe
			ServerSocket serverSocket = new ServerSocket(8039);
			System.out.println("Server open!");
			while (true)// Cho client ket noi
			{
				// Su dung multithread
				// Khi co 1 client gui yeu cau toi thi se tao ra 1 thread phuc vu client do
				new ThreadSocket(serverSocket.accept()).start();

			}
		} catch (IOException e) {
			System.out.println("Exception: " + e.getMessage());

		}
	}
}
