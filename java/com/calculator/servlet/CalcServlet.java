package com.calculator.servlet;

import jakarta.servlet.ServletException;
import jakarta.servlet.http.*;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CalcServlet extends HttpServlet {

    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {

        String num1Str = request.getParameter("num1");
        String num2Str = request.getParameter("num2");
        String type = request.getParameter("type");

        double result = 0.0;
        boolean error = false;
        String operationStr = "";

        try {
            double num1 = Double.parseDouble(num1Str);
            double num2 = Double.parseDouble(num2Str);

            switch (type) {
                case "sum":
                    result = num1 + num2;
                    operationStr = num1 + " + " + num2 + " = " + result;
                    break;
                case "sub":
                    result = num1 - num2;
                    operationStr = num1 + " - " + num2 + " = " + result;
                    break;
                case "mul":
                    result = num1 * num2;
                    operationStr = num1 + " * " + num2 + " = " + result;
                    break;
                case "div":
                    if (Math.abs(num2) < 1e-9) {
                        operationStr = "Error: Division by zero";
                        error = true;
                    } else {
                        result = num1 / num2;
                        operationStr = num1 + " / " + num2 + " = " + result;
                    }
                    break;

                default:
                    operationStr = "Invalid operation type.";
                    error = true;
            }
        } catch (NumberFormatException e) {
            operationStr = "Invalid input: numbers required.";
            error = true;
        }

        HttpSession session = request.getSession();
        List<String> history = (List<String>) session.getAttribute("history");
        if (history == null) {
            history = new ArrayList<>();
        }
        history.add(operationStr);
        session.setAttribute("history", history);

        request.setAttribute("result", operationStr);
        request.setAttribute("history", history);

        request.getRequestDispatcher("/jsp/calculator.jsp").forward(request, response);
    }
}
