package com.calculator.servlet;

import com.calculator.dao.UserDAO;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.*;

import java.io.IOException;

public class LoginServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        String username = request.getParameter("username");
        String password = request.getParameter("password");

        if (username == null || password == null || username.isBlank() || password.isBlank()) {
            request.setAttribute("error", "Username and password cannot be empty.");
            request.getRequestDispatcher("/jsp/login.jsp").forward(request, response);
            return;
        }

        boolean valid = UserDAO.validate(username, password);
        if (valid) {
            HttpSession session = request.getSession();
            session.setAttribute("user", username);
            response.sendRedirect("jsp/calculator.jsp");
        } else {
            request.setAttribute("error", "Invalid username or password.");
            request.getRequestDispatcher("/jsp/login.jsp").forward(request, response);
        }
    }
}
