pipeline {
	agent any

    environment {
		IMAGE_NAME = "timesheetgen:1.0"
    }

    stages {
		stage('Checkout') {
			steps {
				git 'https://github.com/Sudhir1812/TimesheetGen.git'
            }
        }

        stage('Build') {
			steps {
				sh 'mvn clean package'
            }
        }

        stage('Docker Build') {
			steps {
				sh 'docker build -t $IMAGE_NAME .'
            }
        }

        stage('Deploy to Kubernetes') {
			steps {
				sh 'kubectl apply -f k8s/'
            }
        }
    }
}
