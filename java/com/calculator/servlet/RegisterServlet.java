package com.calculator.servlet;

import com.calculator.dao.UserDAO;
import jakarta.servlet.ServletException;
import jakarta.servlet.http.*;

import java.io.IOException;

public class RegisterServlet extends HttpServlet {

    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        String username = request.getParameter("username");
        String password = request.getParameter("password");

        if (username == null || password == null || username.isBlank() || password.isBlank()) {
            request.setAttribute("error", "Username and password cannot be empty.");
            request.getRequestDispatcher("/jsp/register.jsp").forward(request, response);
            return;
        }

        boolean success = UserDAO.register(username, password);
        if (success) {
            response.sendRedirect("jsp/login.jsp");
        } else {
            request.setAttribute("error", "User already exists.");
            request.getRequestDispatcher("/jsp/register.jsp").forward(request, response);
        }
    }
}
