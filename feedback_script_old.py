import pandas as pd
import aiohttp
import asyncio
import logging
import json
from aiohttp import ClientTimeout

# Настройка логирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Загрузка данных
df = pd.read_excel('skills.xlsx', engine="openpyxl")

# Ограничиваем количество одновременных запросов
semaphore = asyncio.Semaphore(5)  # Максимум 5 одновременных запросов

# Асинхронная функция для запроса к Ollama
async def fetch_feedback(session, prompt):
    timeout = ClientTimeout(total=300)  # Увеличиваем timeout до 300 секунд
    async with semaphore:
        try:
            async with session.post(
                "http://127.0.0.1:11434/api/chat",
                json={"model": "mistral", "messages": [{"role": "user", "content": prompt}]},
                timeout=timeout
            ) as response:
                feedback = ""
                async for line in response.content:
                    if line:
                        try:
                            data = json.loads(line.decode('utf-8'))
                            feedback += data.get("message", {}).get("content", "")
                        except json.JSONDecodeError:
                            logging.error("Ошибка декодирования JSON")
                return feedback
        except asyncio.TimeoutError:
            logging.error(f"Тайм-аут при обработке запроса: {prompt[:50]}...")
            return "Ошибка: превышено время ожидания"
        except Exception as e:
            logging.error(f"Ошибка при запросе к Ollama: {e}")
            return "Ошибка: не удалось получить обратную связь"

# Основная асинхронная функция
async def main():
    timeout = ClientTimeout(total=300)  # Увеличиваем timeout до 300 секунд
    async with aiohttp.ClientSession(timeout=timeout) as session:
        tasks = []
        for index, row in df.iterrows():
            # Формируем промпт на основе данных студента
            prompt = (
                f"Составь обратную связь для {row['Имя']}. "
                f"Оценки по компетенциям: "
            )
            for col in df.columns[1:]:  # Перебираем все колонки с компетенциями
                if pd.notna(row[col]):  # Проверяем, что оценка есть
                    prompt += f"{col}: {row[col]}. "
            prompt += (
                "Обратная связь должна быть позитивной, конструктивной и содержать рекомендации для развития. "
                "Учитывай, что не все компетенции оценены."
            )
            tasks.append(fetch_feedback(session, prompt))
            logging.info(f"Обработка студента {index + 1}/{len(df)}: {row['Имя']}")
        
        feedbacks = await asyncio.gather(*tasks)
        for index, feedback in enumerate(feedbacks):
            df.at[index, 'Обратная связь для родителей'] = feedback

# Запуск асинхронного кода
logging.info("Запуск асинхронной генерации обратной связи...")
asyncio.run(main())

# Сохранение результатов
df.to_excel('skills_feedback_with_comments.xlsx', index=False)
logging.info("Файл 'skills_feedback_with_comments.xlsx' успешно создан!")
