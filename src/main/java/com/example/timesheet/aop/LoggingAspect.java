package com.example.timesheet.aop;

import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Before;
import org.aspectj.lang.annotation.Pointcut;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

@Aspect
@Component
public class LoggingAspect {

    private static final Logger logger = LoggerFactory.getLogger(LoggingAspect.class);

    // Pointcut for all public methods inside service package
    @Pointcut("execution(* com.example.timesheet.service..*(..))")
    public void serviceMethods() {}

    // Before executing any service method
    @Before("serviceMethods()")
    public void logBeforeService(JoinPoint joinPoint) {

        String methodName = joinPoint.getSignature().getName();
        String className = joinPoint.getTarget().getClass().getSimpleName();
        Object[] args = joinPoint.getArgs();

        logger.info("➡️ Executing Service Method: {}.{}()", className, methodName);

        if (args.length > 0) {
            logger.info("   📌 Arguments:");
            for (Object arg : args) {
                logger.info("      → {}", arg);
            }
        } else {
            logger.info("   📌 No Arguments");
        }
    }
}
