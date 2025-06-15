from django.shortcuts import render, redirect
from django.contrib import messages
from django.http import JsonResponse
from .models import Subscriber

def index(request):
    if request.method == 'POST' and request.headers.get('X-Requested-With') == 'XMLHttpRequest':
        email = request.POST.get('email')
        try:
            subscriber, created = Subscriber.objects.get_or_create(
                email=email,
                defaults={'is_active': True}
            )
            if not created:
                return JsonResponse({'status': 'error', 'message': 'This email is already subscribed!'}, status=400)
            return JsonResponse({'status': 'success', 'message': 'Thank you for subscribing!'})
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': 'An error occurred. Please try again.'}, status=500)
    
    return render(request, 'index.html')
