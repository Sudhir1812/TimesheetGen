pipeline {
    agent any

    environment {
        APP_NAME   = "timesheetgen"
        IMAGE_TAG  = "${BUILD_NUMBER}"
        IMAGE_NAME = "${APP_NAME}:${IMAGE_TAG}"
        KUBECONFIG = "C:\\ProgramData\\Jenkins\\.kube\\config"
    }

    options {
        disableConcurrentBuilds()
    }

    stages {

        stage('CI - Build & Test') {
            steps {
                bat 'mvnw.cmd clean test'
            }
        }

        stage('Package') {
            steps {
                bat 'mvnw.cmd clean package -DskipTests'
            }
        }

        stage('Docker Build') {
            when {
                anyOf {
                    branch 'develop'
                    branch 'main'
                }
            }
            steps {
                bat "docker build -t %IMAGE_NAME% ."
            }
        }

        stage('Deploy to TEST') {
            when {
                branch 'develop'
        }
            steps {
                echo 'Deploying to TEST environment'
        // Create namespace if it doesn't exist
        bat 'kubectl get namespace test || kubectl create namespace test'
        // Apply manifests
        bat 'kubectl apply -f k8s/test/ --validate=false'
            }
        }

        stage('Deploy to PROD') {
            when {
                branch 'main'
        }
            steps {
                echo 'Deploying to PRODUCTION environment'
        // Create namespace if it doesn't exist
        bat 'kubectl get namespace prod || kubectl create namespace prod'
        // Apply manifests
        bat 'kubectl apply -f k8s/prod/ --validate=false'
            }
        }
    }

    post {
        success {
            echo "Pipeline successful for branch: ${env.BRANCH_NAME}"
        }
        failure {
            echo "Pipeline FAILED for branch: ${env.BRANCH_NAME}"
        }
        always {
            cleanWs()
        }
    }
}
