package client;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.Socket;

public class Client {
	public static void main(String[] args) throws IOException {
		Socket socket = null;

		socket = new Socket("localhost", 8039);
		System.out.println("Connect success!");

		File file = new File(
				"D:\\Learning\\Lectures\\Fourth-yearStudent_1\\Lập trình mạng\\Danh-sach-can-bo-coi-thi.xlsx");

		byte[] bytes = new byte[16 * 1024];
		InputStream in = new FileInputStream(file);
		OutputStream out = socket.getOutputStream();

		int count;
		while ((count = in.read(bytes)) > 0) {
			out.write(bytes, 0, count);
		}

		out.close();
		in.close();
		socket.close();
	}
}
