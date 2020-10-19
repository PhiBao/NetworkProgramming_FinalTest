package server;

import java.net.Socket;

import model.GetResult;
import model.OfficersPost;
import model.RoomsPost;

import java.io.*;

public class ThreadSocket extends Thread {

	Socket socket = null;

	public ThreadSocket(Socket socket) {
		System.out.println("Call to thread socket. ");
		System.out.println("Socket is connected: " + socket.isConnected());
		System.out.println("Socket address: " + socket.getInetAddress().getHostAddress());
		System.out.println("Socket port: " + socket.getPort());
		this.socket = socket;

	}

	public void run() {
		try {
			String destinationDir = "C:\\Users\\Admin\\Documents\\myCode\\Test\\test.xlsx";
			String resultDir = "C:\\Users\\Admin\\Documents\\myCode\\Test\\result.xlsx";
			
			// Tao luong nhan du lieu tu client
			DataInputStream din = new DataInputStream(socket.getInputStream());
			FileOutputStream dout = new FileOutputStream(new File(destinationDir));
			byte[] bytes = new byte[16 * 1024];

			int count;
			while ((count = din.read(bytes)) > 0) {
				dout.write(bytes, 0, count);
			}

			try {
				// Import data to database
				OfficersPost officersPost = new OfficersPost();
				officersPost.insertListOfficers(destinationDir);
				RoomsPost roomsPost = new RoomsPost();
				roomsPost.insertListRooms(destinationDir);
				GetResult getResult = new GetResult();
				getResult.export(resultDir);

			} catch (Exception e) {
				din.close();
				dout.close();
				socket.close();
				System.out.println("Error: " + e.getMessage());

			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
